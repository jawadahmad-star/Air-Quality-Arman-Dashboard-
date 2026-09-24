# -*- coding: utf-8 -*-
"""
Air Quality Project - dashboard builder
=======================================
Reads the SurveyCTO wide export (CSV and/or Stata .dta), the XLSForm (for value
labels) and the assignment frame (prefill_data.xlsx), computes every aggregate,
encrypts the payload with the dashboard password and writes index.html.

    python build_dashboard.py            build index.html
    python build_dashboard.py --check    validate the data only, write nothing

DATA DROP
    Put the SurveyCTO export in  ..\\Data\\   (any of: *.csv, *.dta - or both).
    * Both present  -> the two files are reconciled row-by-row on KEY; the cleaned
                       .dta wins any conflicting cell (set "prefer" in the config).
    * One present   -> that file is used.
    * None present  -> Data\\_dummy\\ is used and the dashboard is stamped DEMO.

PRIVACY
    Names, phone numbers, addresses, device ids and point GPS are dropped on
    load and never reach the payload. Only aggregates are written, and the whole
    payload is AES-256-GCM encrypted (key = PBKDF2 of the password), so the
    public index.html contains no readable data.

Requires: pandas, numpy, openpyxl, cryptography  (pyreadstat optional, for .dta)
"""
import base64
import json
import math
import os
import re
import secrets
import sys
import zlib
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DATA_DIR = Path(os.environ["AQP_DATA_DIR"]) if os.environ.get("AQP_DATA_DIR") else ROOT / "Data"
TPL = HERE / "dashboard_template.html"
OUT = Path(os.environ["AQP_OUT"]) if os.environ.get("AQP_OUT") else HERE / "index.html"
LOG = HERE / "build_log.txt"
CONFIG = HERE / "dashboard.config.json"

FORM_CANDIDATES = [ROOT / "Instrument" / "air_quality.xlsx", HERE / "air_quality.xlsx"]
FRAME_CANDIDATES = [ROOT / "Instrument" / "Attachments" / "prefill_data.xlsx", HERE / "prefill_data.xlsx"]

CHECK_ONLY = "--check" in sys.argv
DUMP = sys.argv[sys.argv.index("--dump") + 1] if "--dump" in sys.argv else None
LOGLINES = []


def say(msg=""):
    print(msg)
    LOGLINES.append(str(msg))


# ====================================================================== config
if not CONFIG.exists():
    sys.exit("dashboard.config.json is missing. It holds the dashboard password and is not in git - "
             "copy it from the project folder or recreate it (see README.md).")
CFG = json.loads(CONFIG.read_text(encoding="utf-8"))
PASSWORD = CFG["password"]
MIN_DUR = float(CFG.get("min_duration_min", 15))
MAX_DUR = float(CFG.get("max_duration_min", 120))
FIELD_HOURS = tuple(CFG.get("field_hours", [7, 21]))
MIN_CELL = int(CFG.get("small_cell_min", 5))
PREFER = CFG.get("prefer", "dta")
ANON_ENUM = bool(CFG.get("anonymise_enumerators", True))

# columns that identify a person, a household or a device: dropped on load
PII_EXACT = {"fname", "lname", "address", "phone_mobile", "phone_landline", "mobile_number", "name_child",
             "school_child", "deviceid", "subscriberid", "simid", "devicephonenum", "username", "caseid",
             "text_audit_file", "enum_other", "tuition_type", "air_school_which", "long_term_medical_o",
             "dec_maker_o", "gender_o", "primary_source_o", "measure_air_pollution_o", "actions_protect_c_o",
             "travel_school_o", "more_pay_o", "text_audit"}
PII_PREFIX = ("geo_", "gps", "phone", "mobile_number")


# ====================================================================== statistics (no scipy)
def _betacf(a, b, x):
    tiny, qab, qap, qam = 1e-30, a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    d = 1.0 / (d if abs(d) > tiny else tiny)
    h = d
    for m in range(1, 300):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d; d = d if abs(d) > tiny else tiny
        c = 1.0 + aa / c; c = c if abs(c) > tiny else tiny
        d = 1.0 / d; h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d; d = d if abs(d) > tiny else tiny
        c = 1.0 + aa / c; c = c if abs(c) > tiny else tiny
        d = 1.0 / d; delta = d * c; h *= delta
        if abs(delta - 1.0) < 3e-12:
            break
    return h


def betai(a, b, x):
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x))
    if x < (a + 1) / (a + b + 2):
        return bt * _betacf(a, b, x) / a
    return 1 - bt * _betacf(b, a, 1 - x) / b


def gammq(a, x):
    """upper regularised incomplete gamma Q(a, x)"""
    if x <= 0:
        return 1.0
    if x < a + 1:
        ap, s, dl = a, 1.0 / a, 1.0 / a
        for _ in range(500):
            ap += 1; dl *= x / ap; s += dl
            if abs(dl) < abs(s) * 1e-13:
                break
        return 1 - s * math.exp(-x + a * math.log(x) - math.lgamma(a))
    tiny = 1e-30
    b = x + 1 - a; c = 1 / tiny; d = 1 / b; h = d
    for i in range(1, 500):
        an = -i * (i - a); b += 2
        d = an * d + b; d = d if abs(d) > tiny else tiny
        c = b + an / c; c = c if abs(c) > tiny else tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) < 1e-13:
            break
    return math.exp(-x + a * math.log(x) - math.lgamma(a)) * h


def t_pvalue(t, df):
    return betai(df / 2.0, 0.5, df / (df + t * t)) if df > 0 else float("nan")


def f_pvalue(F, d1, d2):
    return betai(d2 / 2.0, d1 / 2.0, d2 / (d2 + d1 * F)) if F >= 0 and d1 > 0 and d2 > 0 else float("nan")


def chi2_pvalue(x, k):
    return gammq(k / 2.0, x / 2.0)


def tcrit(df):
    if df < 1:
        return float("nan")
    lo, hi = 0.0, 200.0
    for _ in range(80):
        mid = (lo + hi) / 2
        (lo, hi) = (mid, hi) if t_pvalue(mid, df) > 0.05 else (lo, mid)
    return (lo + hi) / 2


def mean_ci(x):
    x = np.asarray(pd.Series(x).dropna(), float)
    n = len(x)
    if n == 0:
        return None
    m = float(x.mean())
    if n < 2:
        return {"m": m, "lo": None, "hi": None, "n": n}
    se = float(x.std(ddof=1)) / math.sqrt(n)
    tc = tcrit(n - 1)
    return {"m": m, "lo": m - tc * se, "hi": m + tc * se, "n": n}


def welch(a, b):
    """difference b - a, Welch t-test. returns dict or None"""
    a = np.asarray(pd.Series(a).dropna(), float); b = np.asarray(pd.Series(b).dropna(), float)
    if len(a) < 3 or len(b) < 3:
        return None
    va, vb = a.var(ddof=1) / len(a), b.var(ddof=1) / len(b)
    se = math.sqrt(va + vb)
    diff = float(b.mean() - a.mean())
    if se == 0:
        return {"diff": diff, "lo": diff, "hi": diff, "p": 1.0 if diff == 0 else 0.0, "na": len(a), "nb": len(b)}
    df = (va + vb) ** 2 / ((va ** 2) / (len(a) - 1) + (vb ** 2) / (len(b) - 1))
    tc = tcrit(df)
    return {"diff": diff, "lo": diff - tc * se, "hi": diff + tc * se, "p": t_pvalue(diff / se, df),
            "na": len(a), "nb": len(b)}


def anova_p(groups):
    gs = [np.asarray(pd.Series(g).dropna(), float) for g in groups]
    gs = [g for g in gs if len(g) >= 2]
    if len(gs) < 2:
        return None
    n = sum(len(g) for g in gs); k = len(gs)
    grand = np.concatenate(gs).mean()
    ssb = sum(len(g) * (g.mean() - grand) ** 2 for g in gs)
    ssw = sum(((g - g.mean()) ** 2).sum() for g in gs)
    if ssw == 0 or n - k <= 0:
        return None
    return f_pvalue((ssb / (k - 1)) / (ssw / (n - k)), k - 1, n - k)


def chi2_p(flags_by_group):
    """flags_by_group: list of 0/1 arrays, one per arm"""
    fl = [np.asarray(pd.Series(g).dropna(), float) for g in flags_by_group]
    fl = [g for g in fl if len(g)]
    if len(fl) < 2:
        return None
    tot = sum(len(g) for g in fl); ones = sum(g.sum() for g in fl)
    if ones == 0 or ones == tot:
        return None
    chi = 0.0
    for g in fl:
        e1 = len(g) * ones / tot; e0 = len(g) - e1
        chi += (g.sum() - e1) ** 2 / e1 + ((len(g) - g.sum()) - e0) ** 2 / e0
    return chi2_pvalue(chi, len(fl) - 1)


# ====================================================================== load
def _norm_col(c):
    return re.sub(r"[^0-9A-Za-z]", "_", str(c))


def _stringify(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s.dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
    if pd.api.types.is_numeric_dtype(s):
        f = s.astype("float64")
        return f.map(lambda v: "" if pd.isna(v) else (str(int(v)) if float(v).is_integer() and abs(v) < 1e15 else repr(float(v))))
    return s.astype("string").fillna("").str.strip().astype(object)


def load_file(path):
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig", low_memory=False)
    else:
        try:
            import pyreadstat
            df, _ = pyreadstat.read_dta(str(path), apply_value_formats=False)
        except ImportError:
            df = pd.read_stata(path, convert_categoricals=False)
    df.columns = [_norm_col(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    out = pd.DataFrame({c: _stringify(df[c]) for c in df.columns})
    for c in ("starttime", "endtime", "SubmissionDate"):
        if c in out:
            out[c] = pd.to_datetime(out[c], errors="coerce", format="mixed").dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
    out["_key"] = _make_key(out)
    return out


def _make_key(df):
    if "KEY" in df and (df["KEY"] != "").any():
        return df["KEY"]
    g = lambda c: df[c] if c in df else pd.Series("", index=df.index)
    return g("hh_id") + "|" + g("starttime") + "|" + g("enum_name")


def find_sources():
    files = []
    if DATA_DIR.exists():
        for p in DATA_DIR.rglob("*"):
            if p.suffix.lower() not in (".csv", ".dta") or "_dummy" in p.parts:
                continue
            nm = p.name.lower()
            if nm.startswith("ta_") or "prefill" in nm or "check" in nm or "master" in nm:
                continue
            files.append(p)
    csvs = sorted([p for p in files if p.suffix.lower() == ".csv"], key=lambda p: p.stat().st_mtime)
    dtas = sorted([p for p in files if p.suffix.lower() == ".dta"], key=lambda p: p.stat().st_mtime)
    if csvs or dtas:
        return "live", (csvs[-1] if csvs else None), (dtas[-1] if dtas else None)
    dm = DATA_DIR / "_dummy"
    return "demo", next(iter(sorted(dm.glob("*.csv"))), None), next(iter(sorted(dm.glob("*.dta"))), None)


COMPARE_COLS = ["status_survey", "consent_q", "enum_name", "hh_id", "ap_lm_arm", "order_key", "contribute_will",
                "change_pay_w", "age", "hh_income", "area_class", "gender", "hh_member", "duration", "tp_bin"]


def _num_eq(a, b):
    na, nb = pd.to_numeric(a, errors="coerce"), pd.to_numeric(b, errors="coerce")
    both = na.notna() & nb.notna()
    close = (na - nb).abs() < 1e-6
    txt = (a == b)
    return np.where(both, close, txt)


def reconcile(csv_df, dta_df):
    rep = {"csv_rows": len(csv_df), "dta_rows": len(dta_df)}
    ck, dk = set(csv_df["_key"]), set(dta_df["_key"])
    rep["matched"] = len(ck & dk); rep["only_csv"] = len(ck - dk); rep["only_dta"] = len(dk - ck)
    a = csv_df.drop_duplicates("_key").set_index("_key")
    b = dta_df.drop_duplicates("_key").set_index("_key")
    common = a.index.intersection(b.index)
    conflicts, cols_bad = 0, {}
    for c in [c for c in COMPARE_COLS if c in a and c in b]:
        eq = _num_eq(a.loc[common, c].fillna(""), b.loc[common, c].fillna(""))
        bad = int((~eq).sum())
        if bad:
            conflicts += bad; cols_bad[c] = bad
    rep["conflict_cells"] = conflicts; rep["conflict_cols"] = cols_bad
    first, second = (dta_df, csv_df) if PREFER == "dta" else (csv_df, dta_df)
    merged = pd.concat([first, second[~second["_key"].isin(first["_key"])]], ignore_index=True).fillna("")
    rep["prefer"] = PREFER; rep["merged_rows"] = len(merged)
    return merged, rep


mode, csv_path, dta_path = find_sources()
if not csv_path and not dta_path:
    sys.exit("No data found in ..\\Data and no dummy data in ..\\Data\\_dummy. "
             "Run Data\\_dummy\\make_dummy_data.py or drop a SurveyCTO export in ..\\Data.")
say(f"Data mode : {mode.upper()}")
recon = {"mode": mode, "files": [p.name for p in (csv_path, dta_path) if p]}
csv_df = load_file(csv_path) if csv_path else None
dta_df = load_file(dta_path) if dta_path else None
if csv_df is not None and dta_df is not None:
    D, rep = reconcile(csv_df, dta_df)
    recon.update(rep)
    say(f"CSV rows  : {rep['csv_rows']}   DTA rows: {rep['dta_rows']}   matched: {rep['matched']}   "
        f"only-CSV: {rep['only_csv']}   only-DTA: {rep['only_dta']}")
    say(f"Conflicts : {rep['conflict_cells']} cell(s) {rep['conflict_cols'] or ''}  -> '{PREFER}' values used")
    recon["state"] = "both"
elif csv_df is not None:
    D = csv_df; recon.update({"csv_rows": len(D), "state": "csv"})
    say(f"CSV only  : {len(D)} rows ({csv_path.name}) - add the .dta to cross-check")
else:
    D = dta_df; recon.update({"dta_rows": len(D), "state": "dta"})
    say(f"DTA only  : {len(D)} rows ({dta_path.name}) - add the .csv to cross-check")
D = D.reset_index(drop=True).copy()

# ---------------------------------------------------------------- form labels
form_path = next((p for p in FORM_CANDIDATES if p.exists()), None)
if not form_path:
    sys.exit("XLSForm not found (looked in Instrument\\ and Dashboard\\): needed for value labels.")
_ch = pd.read_excel(form_path, sheet_name="choices").fillna("")
_labcol = "label: eng" if "label: eng" in _ch.columns else next(c for c in _ch.columns if str(c).startswith("label"))
CH = {}
for _, _r in _ch.iterrows():
    ln = str(_r["list_name"]).strip()
    if not ln:
        continue
    try:
        v = int(float(_r["value"]))
    except (TypeError, ValueError):
        continue
    lab = re.sub(r"\s+", " ", str(_r[_labcol]).replace("\xa0", " ")).strip()
    if lab:
        CH.setdefault(ln, []).append((v, lab))

SHORT = {
    ("source", 1): "TV news channel", ("source", 2): "Social media", ("source", 3): "News websites",
    ("source", 4): "Mobile apps", ("source", 5): "Newspapers", ("source", 8): "Not applicable",
    ("belief", 1): "Almost no indoor pollution", ("belief", 2): "Indoors more polluted than outdoors",
    ("belief", 3): "A little (much less than outdoors)", ("belief", 4): "Moderately (about half of outdoors)",
    ("belief", 5): "A great deal (almost as bad as outdoors)", ("belief", 6): "Equally polluted",
    ("long_term_med", 3): "Heart / cardiovascular / hypertension", ("long_term_med", 2): "Other lung or breathing problem",
    ("long_term_med", 777): "Other", ("disease_risk", 5): "Doubling (100% increase)",
    ("measures", 777): "Other", ("travel", 777): "Other", ("reason", 777): "Other", ("source", 777): "Other",
    ("support", 1): "Yes, regardless of cost", ("support", 2): "Yes, only if cost increase is small",
    ("support", 3): "No, not if it raises taxes", ("support", 4): "Unsure",
    ("ngo", 2): "PAQI (Pakistan Air Quality Initiative)", ("effects_purifiers", 1): "Totally ineffective",
    ("bid_use", 1): "My first bid", ("bid_use", 2): "My second bid", ("bid_use", 3): "The higher of the two",
    ("yesno_sure", 1): "Yes, still 100%", ("yesno_sure", 2): "No, likely lower in November", ("yesno_sure", 3): "Not sure",
    ("score", 5): "More than 20% higher", ("network_pk", 777): "Other",
}
for (ln, v), lab in SHORT.items():
    CH[ln] = [(vv, lab if vv == v else ll) for vv, ll in CH.get(ln, [])]


def lab_of(ln, code):
    return next((l for v, l in CH.get(ln, []) if v == code), str(code))


# ---------------------------------------------------------------- frame
frame = None
frame_path = next((p for p in FRAME_CANDIDATES if p.exists()), None)
if frame_path:
    _f = pd.read_excel(frame_path, usecols=[0, 1, 2]).dropna()
    _f.columns = ["hh_id", "order_key", "ap_lm_arm"]
    frame = _f.astype(int).drop_duplicates("hh_id").set_index("hh_id")
    say(f"Frame     : {len(frame)} households ({frame_path.name})")
else:
    say("Frame     : none found - targets derived from the data")


# ====================================================================== prepare analysis columns
def num(df, col):
    if col in df:
        return pd.to_numeric(df[col].replace("", np.nan), errors="coerce")
    return pd.Series(np.nan, index=df.index, dtype="float64")


def has(df, col):
    return (df[col] != "") if col in df else pd.Series(False, index=df.index)


D["_dt"] = pd.to_datetime(D.get("starttime", ""), errors="coerce", format="mixed")
D["_end"] = pd.to_datetime(D.get("endtime", ""), errors="coerce", format="mixed")
D["_date"] = D["_dt"].dt.normalize()
D["_status"] = num(D, "status_survey")
D["_hh"] = num(D, "hh_id")
dur = num(D, "duration") / 60.0
alt = (D["_end"] - D["_dt"]).dt.total_seconds() / 60.0
D["_min"] = dur.where(dur.notna(), alt)
D["_arm"] = num(D, "ap_lm_arm")
D["_ord"] = num(D, "order_key")
if frame is not None:
    D["_arm"] = D["_arm"].fillna(D["_hh"].map(frame["ap_lm_arm"]))
    D["_ord"] = D["_ord"].fillna(D["_hh"].map(frame["order_key"]))
D["_enum"] = num(D, "enum_name")
_latcols = [c for c in D.columns if re.match(r"(?i)^geo_2.*lat", c)]
_loncols = [c for c in D.columns if re.match(r"(?i)^geo_2.*lon", c)]
D["_gps"] = False
if _latcols and _loncols:
    D["_gps"] = (pd.to_numeric(D[_latcols[0]].replace("", np.nan), errors="coerce").notna() &
                 pd.to_numeric(D[_loncols[0]].replace("", np.nan), errors="coerce").notna())
D["_gps_col"] = bool(_latcols)
D = D.copy()

# ---- PII guard: drop, then assert
_drop = [c for c in D.columns if (c in PII_EXACT or c.lower().startswith(PII_PREFIX)) and not c.startswith("_")]
D = D.drop(columns=_drop)
assert not [c for c in D.columns if c in PII_EXACT], "PII column survived the guard"

if len(D) and D["_dt"].isna().all():
    say("WARNING: starttime could not be parsed - dates will be empty")

ENUM_NAMES = {v: l for v, l in CH.get("enum_name", [])}


def enum_label(code):
    if pd.isna(code):
        return "Unknown"
    code = int(code)
    return f"Enum {code}" if ANON_ENUM else ENUM_NAMES.get(code, f"Enum {code}")


D["_enum_lab"] = D["_enum"].map(enum_label)

# ---- analysis base: completed interviews, in frame, one per household
D["_complete"] = D["_status"] == 1
comp_all = D[D["_complete"]].copy()
n_out_of_frame = 0
if frame is not None:
    infr = comp_all["_hh"].isin(frame.index)
    n_out_of_frame = int((~infr).sum())
    comp_all = comp_all[infr]
n_dup_rows = int(comp_all.duplicated("_hh", keep="last").sum())
C = comp_all.sort_values("_dt").drop_duplicates("_hh", keep="last").reset_index(drop=True)

N_TARGET = int(CFG.get("target_override") or (len(frame) if frame is not None else max(len(C), 1)))
N_ATT = len(D)
N_C = len(C)
say(f"Attempts  : {N_ATT}   completed (analysis base): {N_C}   target: {N_TARGET}")
say(f"Excluded  : {n_out_of_frame} completed row(s) outside the frame, {n_dup_rows} duplicate household row(s)")

# ---- derived analysis columns on C
C["_bid0"] = num(C, "contribute_will")
_chg = (num(C, "change_pay") == 1) & num(C, "change_pay_w").notna()
C["_bid"] = np.where(_chg, num(C, "change_pay_w"), C["_bid0"])
C["_bid"] = pd.to_numeric(C["_bid"], errors="coerce")
C["_bonus"] = num(C, "contribute_will_bonus")
C["_rand"] = num(C, "main_contribute_rand_number")
C["_paid"] = np.where(C["_rand"].notna() & C["_bid"].notna(), np.where(C["_rand"] <= C["_bid"], C["_rand"], 0), np.nan)
C["_hit"] = np.where(C["_rand"].notna() & C["_bid"].notna(), (C["_rand"] <= C["_bid"]).astype(float), np.nan)
C["_income"] = num(C, "hh_income")
C["_age"] = num(C, "age")
C["_hhsize"] = num(C, "hh_member")

_gr = C["grade_child"].astype(str).str.extract(r"(\d+)")[0] if "grade_child" in C else pd.Series(np.nan, index=C.index)
C["_grade"] = pd.to_numeric(_gr, errors="coerce")
_ex = C["grade_exam"].astype(str).str.extract(r"(\d+(?:\.\d+)?)")[0] if "grade_exam" in C else pd.Series(np.nan, index=C.index)
C["_exam"] = pd.to_numeric(_ex, errors="coerce").where(lambda s: (s >= 0) & (s <= 100))

# implied annual discount rate from the bisection end-node
def _tp_amt(node):
    node = float(np.clip(node, 1, 31))
    return round(2000 * 1.05 * math.pow(3.809524, (node - 1) / 30) / 50) * 50


C["_tpamt"] = num(C, "tp_bin").map(lambda v: np.nan if pd.isna(v) else _tp_amt(v))
C["_disc"] = (C["_tpamt"] / 2000.0 - 1.0) * 100.0
if "tp_valid" in C:
    C.loc[num(C, "tp_valid") == 0, "_disc"] = np.nan

# ====================================================================== segments
SEG_DEFS = [("all", "All respondents", "All", lambda d: pd.Series(True, index=d.index))]
for v, l in [(0, "Control"), (1, "Video 1"), (2, "Video 2")]:
    SEG_DEFS.append((f"arm{v}", l, "Study arm", (lambda vv: lambda d: d["_arm"] == vv)(v)))
for v, l in CH.get("area", []):
    SEG_DEFS.append((f"area{v}", l, "Area", (lambda vv: lambda d: num(d, "area_class") == vv)(v)))
for v, l in CH.get("gender", []):
    if v in (1, 2):
        SEG_DEFS.append((f"g{v}", l, "Respondent gender", (lambda vv: lambda d: num(d, "gender") == vv)(v)))

SEGS, SEG_META = {}, []
for key, label, group, fn in SEG_DEFS:
    sub = C[fn(C)]
    ok = len(sub) >= (1 if key == "all" else MIN_CELL)
    SEG_META.append({"key": key, "label": label, "group": group, "n": int(len(sub)), "ok": bool(ok)})
    if ok:
        SEGS[key] = sub


# ====================================================================== helpers
def r1(x, d=1):
    return None if x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))) else round(float(x), d)


def pct(n, base, d=1):
    return round(100.0 * n / base, d) if base else 0.0


def fmt_n(x):
    return "—" if x is None or (isinstance(x, float) and math.isnan(x)) else f"{x:,.0f}"


def fmt_rs(x):
    return "—" if x is None or (isinstance(x, float) and math.isnan(x)) else f"Rs {x:,.0f}"


def fmt_p(x):
    return "—" if x is None else (f"{x:.0f}%" if x >= 10 or x == 0 else f"{x:.1f}%")


def med(s):
    s = pd.Series(s).dropna()
    return float(s.median()) if len(s) else None


def avg(s):
    s = pd.Series(s).dropna()
    return float(s.mean()) if len(s) else None


def sh(s, cond):
    """share (%) of non-missing values meeting cond"""
    s = pd.Series(s).dropna()
    return pct(int(cond(s).sum()), len(s)) if len(s) else None


def fdate(ts):
    return ts.strftime("%b %d").replace(" 0", " ")


def K(icon, value, label, delta=None, dcls="navy", cls=""):
    return {"i": icon, "v": value, "l": label, "d": delta, "dc": dcls, "c": cls}


def SD(fn):
    """evaluate fn(df) for every visible segment -> segmented payload"""
    out = {}
    for key, sub in SEGS.items():
        try:
            out[key] = fn(sub)
        except Exception as e:                              # never let one card kill the build
            say(f"  ! card error in segment {key}: {e!r}")
            out[key] = None
    return {"_s": out}


def card(kind, cid, title, desc, d, var=None, **kw):
    c = {"id": cid, "kind": kind, "title": title, "desc": desc, "d": d}
    if var:
        c["var"] = var
    c.update(kw)
    return c


def items(labels, counts, base):
    return [{"l": l, "n": int(n), "p": pct(n, base)} for l, n in zip(labels, counts)]


def f_dist(col, lst, hide=(), drop_zero=False):
    def fn(df):
        s = num(df, col).dropna().astype(int)
        base = len(s)
        if not base:
            return None
        its = [{"l": l, "n": int((s == v).sum()), "p": pct(int((s == v).sum()), base)}
               for v, l in CH[lst] if v not in hide]
        if drop_zero:
            its = [i for i in its if i["n"]]
        return {"items": its, "base": base}
    return fn


def _tokens(df, col, codes):
    """per-row sets of selected codes from a select_multiple column (string or dummy columns)"""
    if col in df:
        ans = df[col] != ""
        toks = df[col].where(ans, "").map(lambda t: {int(x) for x in re.findall(r"-?\d+", t)})
        return toks, ans
    dums = {c: f"{col}_{c}" for c in codes if f"{col}_{c}" in df}
    if not dums:
        return pd.Series([set()] * len(df), index=df.index), pd.Series(False, index=df.index)
    m = pd.DataFrame({c: num(df, n) == 1 for c, n in dums.items()})
    return m.apply(lambda r: {c for c, v in r.items() if v}, axis=1), pd.Series(True, index=df.index)


def f_multi(col, lst, hide=(), top=None, sort=True):
    def fn(df):
        codes = [v for v, _ in CH[lst]]
        toks, ans = _tokens(df, col, codes)
        base = int(ans.sum())
        if not base:
            return None
        its = [{"l": l, "n": int(toks[ans].map(lambda t, v=v: v in t).sum())} for v, l in CH[lst] if v not in hide]
        for i in its:
            i["p"] = pct(i["n"], base)
        its = [i for i in its if i["n"]]
        if sort:
            its.sort(key=lambda i: -i["n"])
        if top:
            its = its[:top]
        return {"items": its, "base": base}
    return fn


def f_hist(getter, edges, labels):
    def fn(df):
        s = pd.Series(getter(df)).dropna()
        base = len(s)
        if not base:
            return None
        cats = pd.cut(s, bins=edges, labels=False, right=False, include_lowest=True)
        cnt = [int((cats == i).sum()) for i in range(len(labels))]
        return {"items": items(labels, cnt, base), "base": base}
    return fn


def f_yes(pairs, yes=1, valid=(1, 2)):
    """[(label, col)] -> hbar of % answering yes"""
    def fn(df):
        its, bmax = [], 0
        for l, c in pairs:
            s = num(df, c).dropna()
            s = s[s.isin(valid)]
            if len(s):
                n = int((s == yes).sum()); its.append({"l": l, "n": n, "p": pct(n, len(s))}); bmax = max(bmax, len(s))
        return {"items": its, "base": bmax} if its else None
    return fn


def f_likert(rows, lst):
    def fn(df):
        cats = CH[lst]
        out = []
        for l, c in rows:
            s = num(df, c).dropna().astype(int)
            if len(s):
                out.append({"l": l, "p": [pct(int((s == v).sum()), len(s)) for v, _ in cats], "n": len(s)})
        return {"cats": [l for _, l in cats], "rows": out} if out else None
    return fn


def top2(s, codes):
    s = pd.Series(s).dropna()
    return pct(int(s.isin(codes).sum()), len(s)) if len(s) else None


PRICES = list(range(0, 2001, 100))


def demand(bids):
    b = pd.Series(bids).dropna()
    if not len(b):
        return None
    return [pct(int((b >= p).sum()), len(b)) for p in PRICES]


def thr_price(bids, share=0.30):
    b = np.sort(np.asarray(pd.Series(bids).dropna(), float))[::-1]
    if not len(b):
        return None
    k = max(int(math.ceil(share * len(b))), 1)
    return float(b[k - 1])


# ====================================================================== PANEL 1 - OVERVIEW
PANELS = []
a = D.copy()
dated = a[a["_date"].notna()]
field_dates = sorted(dated["_date"].unique())
n_field_days = len(field_dates)
last_ts = pd.Timestamp(field_dates[-1]) if n_field_days else None
first_ts = pd.Timestamp(field_dates[0]) if n_field_days else None
n_partial = int((D["_status"] == 2).sum())
n_refused = int(D["_status"].isin([3, 4, 5, 6]).sum())
n_notreached = int(D["_status"].isin([7, 8, 10]).sum())
n_inelig = int((D["_status"] == 9).sum())
hh_reached = int(D.loc[D["_hh"].isin(frame.index), "_hh"].nunique()) if frame is not None else int(D["_hh"].nunique())
med_dur = med(C["_min"])
arm_target = {v: int((frame["ap_lm_arm"] == v).sum()) for v in (0, 1, 2)} if frame is not None else {0: 0, 1: 0, 2: 0}
arm_done = {v: int((C["_arm"] == v).sum()) for v in (0, 1, 2)}
if frame is None:
    arm_target = {v: max(arm_done[v], 1) for v in arm_done}
per_day = N_C / n_field_days if n_field_days else 0
remaining = max(N_TARGET - N_C, 0)
days_left = math.ceil(remaining / per_day) if per_day and remaining else 0


def next_workdays(start, k):
    out, d = [], start
    while len(out) < k:
        d = d + timedelta(days=1)
        if d.weekday() != 6:
            out.append(d)
    return out


proj_finish = next_workdays(last_ts, days_left)[-1] if (n_field_days and days_left) else last_ts
conv = pct(N_C, N_ATT)

META = {
    "title": CFG["title"], "subtitle": CFG["subtitle"], "location": CFG["location"],
    "mode": mode, "built": datetime.now().strftime("%d %b %Y, %H:%M"), "built_iso": datetime.now().isoformat(timespec="minutes"),
    "n_complete": N_C, "target": N_TARGET, "pct_complete": pct(N_C, N_TARGET, 0), "n_attempts": N_ATT,
    "field_days": n_field_days, "last_date": fdate(last_ts) if last_ts is not None else "—",
    "first_date": fdate(first_ts) if first_ts is not None else "—",
    "domain": CFG.get("domain", ""), "recon": recon,
    "arm_done": arm_done, "arm_target": arm_target, "median_min": r1(med_dur),
    "n_enums": int(D["_enum"].nunique()), "min_cell": MIN_CELL,
}

# ---- daily series (field days only)
daily = []
for d in field_dates:
    day = D[D["_date"] == d]
    daily.append({"d": fdate(pd.Timestamp(d)), "c": int((C["_date"] == d).sum()), "a": int(len(day))})
_vals = [x["c"] for x in daily]
_roll = [round(float(np.mean(_vals[max(0, i - 2):i + 1])), 1) for i in range(len(_vals))]

# ---- funnel
n_consent = int((num(D, "consent_q") == 1).sum())
n_bid = int(C["_bid"].notna().sum())
n_topup = int(has(C, "mobile_network").sum() or num(C, "mobile_network").notna().sum())
funnel = [
    {"l": "Fieldwork attempts", "n": N_ATT},
    {"l": "Reached and consented", "n": n_consent},
    {"l": "Completed interviews", "n": int((D["_status"] == 1).sum())},
    {"l": "Counted in analysis (unique, in frame)", "n": N_C},
    {"l": "Bid recorded (WTP exercise)", "n": n_bid},
]

# ---- disposition
disp_codes = CH.get("status_survey", [])
disp = [{"l": l, "n": int((D["_status"] == v).sum())} for v, l in disp_codes]
disp = [x for x in disp if x["n"]]
for x in disp:
    x["p"] = pct(x["n"], N_ATT)

# ---- assignment blocks (arm x order)
blocks = []
ARM_LAB = {0: "Control", 1: "Video 1", 2: "Video 2"}
ORD_LAB = {1: "beliefs asked before bid", 2: "beliefs asked after bid"}
for arm in (0, 1, 2):
    for o in (1, 2):
        tgt = int(((frame["ap_lm_arm"] == arm) & (frame["order_key"] == o)).sum()) if frame is not None else 0
        dn = int(((C["_arm"] == arm) & (C["_ord"] == o)).sum())
        if tgt == 0:
            tgt = max(dn, 1)
        p_ = pct(dn, tgt, 0)
        blocks.append({"bin": f"{ARM_LAB[arm]} · order {o}", "sub": ORD_LAB[o], "done": dn, "target": tgt, "pct": p_,
                       "status": "Completed" if dn >= tgt else ("In Progress" if dn else "Not Started"), "arm": arm})

# ---- household coverage by area
wtp_all = C["_bid"]
wtp_mean = avg(wtp_all)
share_pos = sh(wtp_all, lambda s: s > 0)

ov_kpis = [
    K("✅", fmt_n(N_C), "Completed Interviews", f"{META['pct_complete']:.0f}% of {N_TARGET:,} target", "up", "green"),
    K("🎯", fmt_n(remaining), "Households Remaining", f"{arm_target[0]} control · {arm_target[1]} video 1 · {arm_target[2]} video 2 targeted", "navy", "navy"),
    K("🚪", fmt_n(N_ATT), "Fieldwork Attempts", f"{conv:.0f}% ended in a completed interview", "neutral", "teal"),
    K("🏠", f"{hh_reached} / {len(frame) if frame is not None else N_TARGET}", "Households Reached", "At least one visit recorded", "navy", "purple"),
    K("⏱️", r1(med_dur) if med_dur is not None else "—", "Median Interview (min)", f"Across {n_field_days} field days", "up", "amber"),
    K("💰", fmt_rs(wtp_mean), "Mean Willingness to Pay", f"{fmt_p(share_pos)} bid above zero", "up", ""),
]
ov_callouts = [
    {"cls": "navy", "icon": "⚪", "h": "Control — no video", "p": f"{arm_done[0]} of {arm_target[0]} completed ({pct(arm_done[0], arm_target[0], 0):.0f}%)."},
    {"cls": "teal", "icon": "🎬", "h": "Video 1 — information film", "p": f"{arm_done[1]} of {arm_target[1]} completed ({pct(arm_done[1], arm_target[1], 0):.0f}%)."},
    {"cls": "purple", "icon": "🎞️", "h": "Video 2 — information film", "p": f"{arm_done[2]} of {arm_target[2]} completed ({pct(arm_done[2], arm_target[2], 0):.0f}%)."},
    {"cls": "amber", "icon": "📉", "h": "Non-response", "p": f"{n_refused} refusals and {n_notreached} not reached out of {N_ATT} attempts ({pct(n_refused + n_notreached, N_ATT, 0):.0f}%)."},
]

# ---- key findings (auto-written from the data)
kf = []
if N_C:
    kf.append({"cls": "primary", "html": f"<strong>Willingness to pay:</strong> the average parent would contribute <strong>{fmt_rs(wtp_mean)}</strong> "
               f"(median {fmt_rs(med(wtp_all))}); <strong>{fmt_p(share_pos)}</strong> bid above zero and "
               f"<strong>{fmt_p(sh(wtp_all, lambda s: s >= 2000))}</strong> bid the Rs&nbsp;2,000 maximum. "
               f"The classroom needs 30% of parents to contribute: that threshold is cleared up to a price of <strong>{fmt_rs(thr_price(wtp_all))}</strong>."})
    aq = avg(num(C, "outdoor_air_quality"))
    aqi_ok = f_multi("beleifs_air_pollute", "belief_air")(C)
    aqi_share = next((i["p"] for i in aqi_ok["items"] if i["l"].lower().startswith("levels")), None) if aqi_ok else None
    kf.append({"cls": "navy", "html": f"<strong>Awareness:</strong> parents rate outdoor air quality <strong>{r1(aq) if aq is not None else '—'}/10</strong> on the pollution scale, "
               f"yet only <strong>{fmt_p(aqi_share)}</strong> connect the AQI with levels of air pollution."})
arms_ok = all(int((C["_arm"] == v).sum()) >= 8 for v in (0, 1, 2)) if N_C else False
if arms_ok:
    m0, m1, m2 = (avg(C.loc[C["_arm"] == v, "_bid"]) for v in (0, 1, 2))
    w1, w2 = welch(C.loc[C["_arm"] == 0, "_bid"], C.loc[C["_arm"] == 1, "_bid"]), welch(C.loc[C["_arm"] == 0, "_bid"], C.loc[C["_arm"] == 2, "_bid"])
    kf.append({"cls": "purple", "html": f"<strong>First read on the videos:</strong> mean bid is <strong>{fmt_rs(m0)}</strong> in Control, <strong>{fmt_rs(m1)}</strong> after Video 1 "
               f"and <strong>{fmt_rs(m2)}</strong> after Video 2 (p = {w1['p']:.2f} and {w2['p']:.2f} against Control). Treat as indicative until the sample is complete."})
kf.append({"cls": "teal", "html": f"<strong>Pace:</strong> {per_day:.1f} completed interviews per field day; "
           f"<strong>{fmt_n(remaining)}</strong> households remain" + (f", so the target is reached around <strong>{fdate(proj_finish)}</strong> at this rate." if days_left else ".")})

overview = {
    "id": "overview", "tab": "📊 Overview", "eyebrow": "Section 01", "title": "Programme at a Glance",
    "blurb": f"Where fieldwork stands against the {N_TARGET:,}-household sample: completed interviews, how the three randomised study arms are filling, "
             "what happened at every attempt, and the headline findings so far.",
    "seg": False, "blocks": [
        {"t": "kpis", "d": ov_kpis},
        {"t": "callouts", "d": ov_callouts},
        {"t": "insights", "d": kf, "row": False},
        {"t": "grid", "cols": 1, "cards": [
            card("lines", "ovDaily", "Daily Completed Interviews", "Completed interviews per field day with a 3-day rolling average",
                 {"x": [x["d"] for x in daily], "series": [
                     {"name": "Completed", "slot": 2, "y": _vals, "fill": True, "pts": True, "labels": True},
                     {"name": "3-day average", "slot": 1, "y": _roll, "dash": True}]},
                 opt={"yfmt": "n"}, full=True)]},
        {"t": "grid", "cols": 3, "cards": [
            card("gauge", "ovProg", "Overall Progress vs Target", f"Completed interviews against the {N_TARGET:,}-household frame", {"done": N_C, "target": N_TARGET}),
            card("donut", "ovArm", "Completed by Study Arm", "Control · Video 1 · Video 2",
                 {"items": [{"l": ARM_LAB[v], "n": arm_done[v], "p": pct(arm_done[v], N_C)} for v in (0, 1, 2)], "base": N_C}, opt={"colors": [7, 3, 5], "fmt": "n"}),
            card("hbar", "ovDisp", "Attempt Disposition", "Outcome of every fieldwork attempt recorded so far", {"items": disp, "base": N_ATT}, var="status_survey", opt={"color": 2, "fmt": "n"})]},
        {"t": "grid", "cols": 2, "cards": [
            card("funnel", "ovFunnel", "From Attempt to Analysed Interview", "How many attempts survive each stage", {"rows": funnel}),
            card("blocks", "ovBlocks", "Progress by Assignment Block", "Study arm × order in which beliefs about other parents were asked", {"rows": blocks})]},
        {"t": "note", "html": f"Sampling frame: <strong>{N_TARGET:,}</strong> households pre-assigned to study arm and question order (Control = {arm_target[0]}, Video 1 = {arm_target[1]}, Video 2 = {arm_target[2]}). "
                             "Completed = <code>status_survey</code> is “Completed”, counted once per household. Research Solutions (M&amp;A Research Solutions LLC) | www.rs.org.pk"},
    ]}
PANELS.append(overview)

# ====================================================================== PANEL 2 - FIELD OPERATIONS
cum, run = [], 0
for x in daily:
    run += x["c"]; cum.append(run)
proj_x, proj_y = [], []
if days_left and n_field_days:
    pdays = next_workdays(last_ts, min(days_left, 60))
    proj_x = [fdate(d) for d in pdays]
    proj_y = [min(N_TARGET, round(N_C + per_day * (i + 1))) for i in range(len(pdays))]
    if proj_y:
        proj_y[-1] = N_TARGET if len(pdays) == days_left else proj_y[-1]
lx = [x["d"] for x in daily] + proj_x
ser_actual = cum + [None] * len(proj_x)
ser_proj = ([None] * (len(cum) - 1) + [cum[-1]] + proj_y) if cum and proj_x else None
series_cum = [{"name": "Cumulative completed", "slot": 6, "y": ser_actual, "fill": True, "pts": True, "lastlabel": True}]
if ser_proj:
    series_cum.append({"name": "Projected at current pace", "slot": 4, "y": ser_proj, "dash": True})
series_cum.append({"name": "Frame target", "slot": 2, "y": [N_TARGET] * len(lx), "dash": True})

# ---- calendar heatmap (Mon..Sun rows are weeks)
cal = {"weeks": [], "max": max([x["c"] for x in daily] + [1])}
if n_field_days:
    start = first_ts - timedelta(days=first_ts.weekday())
    end = last_ts + timedelta(days=6 - last_ts.weekday())
    dmap = {pd.Timestamp(d): (int((C["_date"] == d).sum()), int((D["_date"] == d).sum())) for d in field_dates}
    cur = start
    while cur <= end:
        wk = {"label": f"Week of {fdate(cur)}", "days": []}
        for i in range(7):
            dd = cur + timedelta(days=i)
            if dd in dmap:
                wk["days"].append({"d": fdate(dd), "c": dmap[dd][0], "a": dmap[dd][1]})
            elif first_ts <= dd <= last_ts:
                wk["days"].append({"d": fdate(dd), "c": 0, "a": 0, "off": True})
            else:
                wk["days"].append(None)
        cal["weeks"].append(wk)
        cur += timedelta(days=7)

# ---- QA checks
qa_rows = []
qa_flag = {}


def qa(name, defn, n, base=None, hard=False):
    st = "ok" if n == 0 else ("check" if not hard else "check")
    qa_rows.append({"name": name, "def": defn, "n": int(n), "base": base, "st": st})


short = C[(C["_min"] < MIN_DUR)]
long_ = C[(C["_min"] > MAX_DUR)]
hrs = C["_dt"].dt.hour
odd = C[(hrs < FIELD_HOURS[0]) | (hrs >= FIELD_HOURS[1])]
qa("Very short interviews", f"Completed interviews under {MIN_DUR:.0f} minutes", len(short), N_C)
qa("Very long interviews", f"Completed interviews over {MAX_DUR:.0f} minutes", len(long_), N_C)
qa("Started outside field hours", f"Interview start before {FIELD_HOURS[0]}:00 or after {FIELD_HOURS[1]}:00", len(odd), N_C)
if D["_gps_col"].iloc[0] if len(D) else False:
    qa("No GPS point captured", "Completed interviews without a usable location", int((~C["_gps"]).sum()), N_C)
qa("Duplicate household submissions", "Households with more than one completed submission (latest kept)", n_dup_rows, len(comp_all))
if frame is not None:
    qa("Household not in assignment frame", "Completed submissions with an ID outside the frame (excluded)", n_out_of_frame, n_out_of_frame + len(comp_all))
qa("Completed but no bid", "Completed interviews with no willingness-to-pay amount", int(C["_bid"].isna().sum()), N_C)
qa("Implausible income", "Monthly household income under Rs 5,000 or over Rs 1,000,000", int(((C["_income"] < 5000) | (C["_income"] > 1_000_000)).sum()), N_C)
n_flag_total = sum(r["n"] for r in qa_rows)

# ---- enumerator scorecard
enum_rows = []
overall_cap = sh(C["_bid"], lambda s: s >= 2000) or 0
for code, grp in D.groupby("_enum", dropna=True):
    lab = enum_label(code)
    cg = C[C["_enum"] == code]
    n_att, n_cmp = len(grp), int((grp["_status"] == 1).sum())
    cap = sh(cg["_bid"], lambda s: s >= 2000)
    flags = []
    if len(cg) and int((cg["_min"] < MIN_DUR).sum()):
        flags.append(f"{int((cg['_min'] < MIN_DUR).sum())} short")
    if len(cg) >= 5 and cap is not None and cap >= 40 and cap >= 2 * max(overall_cap, 1):
        flags.append("bids cluster at Rs 2,000")
    hh_ = cg["_dt"].dt.hour
    if len(cg) and int(((hh_ < FIELD_HOURS[0]) | (hh_ >= FIELD_HOURS[1])).sum()):
        flags.append("odd-hour starts")
    if n_att >= 8 and pct(n_cmp, n_att) < 0.5 * conv:
        flags.append("low conversion")
    enum_rows.append({"e": lab, "att": n_att, "cmp": n_cmp, "conv": {"v": pct(n_cmp, n_att), "t": f"{pct(n_cmp, n_att):.0f}%"},
                      "med": {"v": r1(med(cg["_min"])), "t": "—" if med(cg["_min"]) is None else f"{med(cg['_min']):.0f}"},
                      "mean": {"v": r1(avg(cg["_bid"]), 0), "t": fmt_rs(avg(cg["_bid"]))},
                      "cap": {"v": cap, "t": "—" if cap is None else f"{cap:.0f}%"},
                      "flags": {"v": len(flags), "t": flags}})
enum_rows.sort(key=lambda r: -r["cmp"])

dur_hist = f_hist(lambda d: d["_min"], [0, 15, 25, 35, 45, 60, 90, 1e9], ["<15", "15–24", "25–34", "35–44", "45–59", "60–89", "90+"])(C)
hour_hist = f_hist(lambda d: d["_dt"].dt.hour, list(range(6, 23)), [str(h) for h in range(6, 22)])(C)
enum_bar = {"items": [{"l": r["e"], "n": r["cmp"], "p": pct(r["cmp"], N_C)} for r in enum_rows], "base": N_C}

recon_card = {"rows": [], "state": recon.get("state"), "mode": mode}
for k_, lbl in (("csv_rows", "CSV rows"), ("dta_rows", "DTA rows"), ("matched", "Rows matched on KEY"), ("only_csv", "Only in CSV"),
                ("only_dta", "Only in DTA"), ("conflict_cells", "Cells that differ")):
    if k_ in recon:
        recon_card["rows"].append({"l": lbl, "v": recon[k_]})
recon_card["files"] = recon.get("files", [])
recon_card["prefer"] = PREFER

ops_kpis = [
    K("📅", n_field_days, "Field Days Elapsed", f"Last activity {META['last_date']}", "navy", "navy"),
    K("⚡", r1(per_day), "Avg Completes / Field Day", "Across the field window", "up", "green"),
    K("👥", META["n_enums"], "Active Enumerators", f"~{round(N_C / META['n_enums']) if META['n_enums'] else 0} completes each", "navy", "teal"),
    K("⏱️", r1(med_dur) if med_dur is not None else "—", "Median Duration (min)", "Start to submission", "neutral", "amber"),
    K("🔍", n_flag_total, "Records to Verify", "Across the automated checks below", "down" if n_flag_total else "up", "purple"),
]
ops_ins = [
    {"cls": "teal", "html": f"<strong>Pace:</strong> the team averages <strong>{per_day:.1f}</strong> completed interviews per field day. At that rate the remaining <strong>{remaining:,}</strong> households need roughly <strong>{days_left}</strong> more field days" + (f", finishing around <strong>{fdate(proj_finish)}</strong>." if days_left else ".")},
    {"cls": "amber", "html": f"<strong>Conversion:</strong> <strong>{conv:.0f}%</strong> of attempts end in a completed interview, about <strong>{(N_ATT / N_C):.1f}</strong> attempts per completed case. " + ("Non-response, not enumerator output, is the constraint on the field plan." if conv < 70 else "Contact is efficient; the field plan is limited by pace, not access.") if N_C else "No completed interviews yet."},
]
PANELS.append({
    "id": "ops", "tab": "🛠️ Field Operations", "eyebrow": "Section 02", "title": "Field Operations & Data Quality",
    "blurb": "How the sample is accumulating against a linear pace, how output and interview length vary across the team, and which records the automated checks ask the field team to verify.",
    "seg": False, "blocks": [
        {"t": "kpis", "d": ops_kpis},
        {"t": "grid", "cols": 1, "cards": [card("lines", "opsCum", "Cumulative Completion vs Sampling Target",
                                                 "Running total of completed interviews, projected forward at the average pace observed so far",
                                                 {"x": lx, "series": series_cum, "max": N_TARGET}, opt={"yfmt": "n", "target": N_TARGET}, tall=True, full=True)]},
        {"t": "grid", "cols": 2, "cards": [
            card("calendar", "opsCal", "Field Calendar", "Completed interviews per day: darker cells are busier days; grey cells are days with no visits", cal),
            card("qa", "opsQa", "Records to Verify", "Automated checks: a flag asks the team to look again, it does not mean the record is wrong", {"rows": qa_rows})]},
        {"t": "grid", "cols": 1, "cards": [card("table", "opsEnum", "Enumerator Scorecard", "Attempts, conversion, interview length and bid pattern by field team member. Click a column to sort.",
                                                {"cols": [{"k": "e", "l": "Enumerator"}, {"k": "att", "l": "Attempts", "num": True}, {"k": "cmp", "l": "Completed", "num": True},
                                                          {"k": "conv", "l": "Conversion", "num": True}, {"k": "med", "l": "Median min", "num": True},
                                                          {"k": "mean", "l": "Mean bid", "num": True}, {"k": "cap", "l": "Bids at Rs 2,000", "num": True},
                                                          {"k": "flags", "l": "To verify", "flags": True}], "rows": enum_rows}, full=True, opt={"sortKey": "cmp"})]},
        {"t": "grid", "cols": 3, "cards": [
            card("bar", "opsDur", "Interview Duration", "Completed interviews by elapsed time: a tight distribution is the expected shape", dur_hist, opt={"color": 4, "fmt": "pct"}),
            card("bar", "opsHour", "Interview Start Hour", "Completed interviews by the hour of day they began (24-hour clock)", hour_hist, opt={"color": 1, "fmt": "n"}),
            card("hbar", "opsEnum2", "Output by Enumerator", "Completed interviews per field team member", enum_bar, opt={"color": 1, "fmt": "n"})]},
        {"t": "grid", "cols": 1, "cards": [card("recon", "opsRecon", "Data Source Check", "How the CSV and Stata exports compare before they are combined", recon_card)]},
        {"t": "insights", "d": ops_ins, "row": True},
        {"t": "note", "html": "Enumerators are shown by staff code. Duration is start to submission. Checks are prompts for verification, never findings of misconduct. The projected finish extrapolates the average daily completion rate and skips Sundays."},
    ]})

# ====================================================================== PANEL 3 - HOUSEHOLDS
def profile_kpis(df):
    return [
        K("👩", f"{len(df):,}", "Respondents", f"{fmt_p(sh(num(df, 'gender'), lambda s: s == 2))} female", "navy", "navy"),
        K("🎂", r1(med(df["_age"]), 0) or "—", "Median Age (years)", f"Mean {r1(avg(df['_age'])) if avg(df['_age']) is not None else '—'}", "neutral", "teal"),
        K("💵", fmt_rs(med(df["_income"])), "Median Monthly Income", "Household, all sources", "neutral", "amber"),
        K("👨‍👩‍👧‍👦", r1(avg(df["_hhsize"])) or "—", "Mean Household Size", "Members currently living in the home", "navy", "purple"),
        K("❄️", fmt_p(sh(num(df, "asset_ac"), lambda s: s == 1)), "Own an Air Conditioner", f"{fmt_p(sh(num(df, 'asset_airpurifier'), lambda s: s == 1))} own an air purifier", "up", "green"),
    ]


def profile_ins(df):
    ac = sh(num(df, "asset_ac"), lambda s: s == 1); ap = sh(num(df, "asset_airpurifier"), lambda s: s == 1)
    ill = sh(num(df, "hh_resp_ill"), lambda s: s == 1)
    return [{"cls": "navy", "html": f"<strong>Equipment gap:</strong> <strong>{fmt_p(ac)}</strong> of households own an air conditioner but only <strong>{fmt_p(ap)}</strong> own an air purifier, so most parents are valuing a product they have never used at home."},
            {"cls": "amber", "html": f"<strong>Health exposure:</strong> in <strong>{fmt_p(ill)}</strong> of households someone has a respiratory illness; the median household earns <strong>{fmt_rs(med(df['_income']))}</strong> a month."}]


assets = [("Air conditioner", "asset_ac"), ("Mobile phone", "asset_mobile"), ("Internet", "asset_internet"), ("Air purifier", "asset_airpurifier")]
PANELS.append({
    "id": "profile", "tab": "🏘️ Households", "eyebrow": "Section 03 · Respondent & household profile", "title": "Who We Interviewed",
    "blurb": "The parents and caregivers in the sample: where they live, their age and income, the size of the household and the equipment they own. Use the segment selector to compare study arms, areas or gender.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(profile_kpis)},
        {"t": "grid", "cols": 3, "cards": [
            card("donut", "prArea", "Area classification", "Urban, peri-urban or rural, as observed by the enumerator", SD(f_dist("area_class", "area")), var="area_class", opt={"colors": [4, 3, 6]}),
            card("donut", "prGender", "Respondent gender", "Who answered the survey", SD(f_dist("gender", "gender")), var="gender", opt={"colors": [1, 7, 8]}),
            card("donut", "prDec", "Primary decision-maker on schooling", "Who usually decides about the child's schooling", SD(f_dist("dec_maker", "decision_maker")), var="dec_maker", opt={"colors": [7, 1, 5, 8]})]},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "prAge", "Age of respondent", "Years, grouped", SD(f_hist(lambda d: d["_age"], [18, 30, 40, 50, 60, 200], ["18–29", "30–39", "40–49", "50–59", "60+"])), var="age", opt={"color": 1}),
            card("bar", "prInc", "Monthly household income", "PKR, all household sources", SD(f_hist(lambda d: d["_income"], [0, 30000, 50000, 75000, 100000, 150000, 1e12], ["<30k", "30–49k", "50–74k", "75–99k", "100–149k", "150k+"])), var="hh_income", opt={"color": 4})]},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "prSize", "Household size", "Members currently living in the household", SD(f_hist(lambda d: d["_hhsize"], [0, 3, 5, 7, 9, 100], ["1–2", "3–4", "5–6", "7–8", "9+"])), var="hh_member", opt={"color": 5}),
            card("hbar", "prAssets", "Household assets", "Share owning each item (head of household, last year)", SD(f_yes(assets)), opt={"color": 6, "fmt": "pct"})]},
        {"t": "grid", "cols": 3, "cards": [
            card("donut", "prWin", "Windows open on arrival", "Enumerator observation on entering the home", SD(f_dist("windows_open", "yesno_2")), var="windows_open", opt={"colors": [3, 1, "muted"]}),
            card("donut", "prIll", "Respiratory illness in household", "Anyone in the household", SD(f_dist("hh_resp_ill", "yesno")), var="hh_resp_ill", opt={"colors": [2, "muted"]}),
            card("donut", "prIll2", "Who is affected", "Child or adult, where illness reported", SD(f_dist("resp_ill_child_adult", "person")), var="resp_ill_child_adult", opt={"colors": [5, 1]})]},
        {"t": "insights", "d": SD(profile_ins), "row": True},
        {"t": "note", "html": "Household identifiers, names, phone numbers, addresses and GPS points are never loaded into this dashboard. Income is self-reported and grouped for display."},
    ]})

# ====================================================================== PANEL 4 - CHILD, SCHOOL & HEALTH
def child_kpis(df):
    lt, ans = _tokens(df, "long_term_medical", [v for v, _ in CH["long_term_med"]])
    anyc = pct(int(lt[ans].map(lambda t: bool(t) and 7 not in t).sum()), int(ans.sum())) if ans.sum() else None
    return [
        K("🩺", fmt_p(anyc), "Child with a Long-term Condition", "Past two months, any listed condition", "down", "red" if False else ""),
        K("📆", r1(avg(num(df, "unable_to_school"))) or "—", "Mean Days Unable to Attend", "Among children reporting a condition", "neutral", "amber"),
        K("🌫️", fmt_p(sh(num(df, "sickness_air"), lambda s: s == 1)), "Link Illness to Air Pollution", "Yes, among those asked", "neutral", "teal"),
        K("🚌", r1(avg(num(df, "time_reach")), 0) or "—", "Mean Commute (min)", "Home to school", "navy", "navy"),
        K("📚", fmt_p(sh(num(df, "tuition"), lambda s: s == 1)), "Attend Tuition / Academy", "After school", "neutral", "purple"),
    ]


def child_ins(df):
    return [{"cls": "amber", "html": f"<strong>Missed school:</strong> children with a reported condition missed an average of <strong>{r1(avg(num(df, 'unable_to_school'))) or '—'}</strong> days in the past two months."},
            {"cls": "teal", "html": f"<strong>Attribution:</strong> <strong>{fmt_p(sh(num(df, 'sickness_air'), lambda s: s == 1))}</strong> of parents link their child's symptoms to air pollution, while <strong>{fmt_p(sh(num(df, 'sickness_air'), lambda s: s == 99))}</strong> say they do not know."}]


PANELS.append({
    "id": "child", "tab": "🎒 Child & School", "eyebrow": "Section 04 · School, commute & child health", "title": "Child, School & Health",
    "blurb": "The child at the centre of the survey: grade and recent exam results, how they get to school, how much schooling illness has cost them, and whether parents connect that illness to air pollution.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(child_kpis)},
        {"t": "grid", "cols": 3, "cards": [
            card("bar", "chGrade", "Child's grade", "Class of the child in the study classroom", SD(f_hist(lambda d: d["_grade"], [1, 6, 7, 8, 9, 10, 13], ["≤5", "6", "7", "8", "9", "10+"])), var="grade_child", opt={"color": 1}),
            card("bar", "chExam", "Most recent exam result", "Percent, mid-term or final, as reported by the parent", SD(f_hist(lambda d: d["_exam"], [0, 50, 60, 70, 80, 90, 101], ["<50", "50–59", "60–69", "70–79", "80–89", "90+"])), var="grade_exam", opt={"color": 6}),
            card("bar", "chTime", "Commute time to school", "Minutes, one way", SD(f_hist(lambda d: num(d, "time_reach"), [0, 10, 20, 30, 45, 1000], ["<10", "10–19", "20–29", "30–44", "45+"])), var="time_reach", opt={"color": 4})]},
        {"t": "grid", "cols": 2, "cards": [
            card("hbar", "chTravel", "How the child travels to school", "Primary mode of travel", SD(f_dist("travel_school", "travel", drop_zero=True)), var="travel_school", opt={"color": 3}),
            card("hbar", "chCond", "Long-term conditions, past two months", "Multiple response: share of children with each condition", SD(f_multi("long_term_medical", "long_term_med")), var="long_term_medical", opt={"color": 2})]},
        {"t": "grid", "cols": 3, "cards": [
            card("bar", "chDays", "Days unable to attend school", "Past two months, among children with a condition", SD(f_hist(lambda d: num(d, "unable_to_school"), [0, 1, 3, 6, 11, 1000], ["0", "1–2", "3–5", "6–10", "11+"])), var="unable_to_school", opt={"color": 2}),
            card("donut", "chAir", "Illness linked to air pollution", "Do parents think symptoms relate to exposure", SD(f_dist("sickness_air", "yesno_dk")), var="sickness_air", opt={"colors": [2, 1, "muted"]}),
            card("donut", "chTut", "Tuition or academy after school", "Child attends", SD(f_dist("tuition", "yesno")), var="tuition", opt={"colors": [5, "muted"]})]},
        {"t": "insights", "d": SD(child_ins), "row": True},
        {"t": "note", "html": "Health items are parent-reported and refer to the two months before the interview. Exam results are shown only where the parent gave a percentage."},
    ]})

# ====================================================================== PANEL 5 - AWARENESS & PERCEPTIONS
def aware_kpis(df):
    aq = f_multi("beleifs_air_pollute", "belief_air")(df)
    aqi = next((i["p"] for i in (aq or {"items": []})["items"] if i["l"].lower().startswith("levels")), None)
    chk = num(df, "check_air_quality").dropna()
    chk = chk[chk.isin([1, 2, 3, 4, 5])]
    bel = num(df, "pollute_believe").dropna(); bel = bel[bel.isin([1, 2, 3, 4, 5, 6])]
    risk = num(df, "risk_air_child").dropna(); risk = risk[risk.isin([1, 2, 3, 4, 5])]
    return [
        K("🔎", fmt_p(aqi), "Know What the AQI Measures", "Chose “levels of air pollution”", "up", "teal"),
        K("🌡️", r1(avg(num(df, "outdoor_air_quality"))) or "—", "Outdoor Pollution Rating (0–10)", "10 = extremely polluted", "down", "amber"),
        K("📱", fmt_p(pct(int(chk.isin([1, 2]).sum()), len(chk)) if len(chk) else None), "Check Air Quality Weekly or More", "Daily or weekly", "neutral", "navy"),
        K("🏠", fmt_p(pct(int(bel.isin([2, 5, 6]).sum()), len(bel)) if len(bel) else None), "Indoor Air ≈ Outdoor or Worse", "Among those with a view", "neutral", "purple"),
        K("⚠️", fmt_p(pct(int(risk.isin([4, 5]).sum()), len(risk)) if len(risk) else None), "See ≥50% Higher Child Risk", "Acute respiratory illness, among those with a view", "down", ""),
    ]


def aware_ins(df):
    ap = top2(num(df, "score_exams_ap"), [4, 5]); ac = top2(num(df, "score_exams_ac"), [4, 5])
    sc = num(df, "air_school").dropna()
    return [{"cls": "purple", "html": f"<strong>Purifier vs. air conditioner:</strong> <strong>{fmt_p(ap)}</strong> of parents expect at least a 10% exam gain from a classroom air purifier, against <strong>{fmt_p(ac)}</strong> for an air conditioner."},
            {"cls": "navy", "html": f"<strong>School action:</strong> only <strong>{fmt_p(pct(int((sc == 1).sum()), len(sc)) if len(sc) else None)}</strong> say their child's school has any air quality measure in place."}]


PANELS.append({
    "id": "aware", "tab": "🌫️ Awareness", "eyebrow": "Section 05 · Beliefs & perceived risk", "title": "Awareness & Perceptions of Air Pollution",
    "blurb": "What parents understand about air pollution and the AQI, where they get information, how polluted they believe the home is, how much risk they see for the child, and which protective measures they know about or use.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(aware_kpis)},
        {"t": "grid", "cols": 3, "cards": [
            card("hbar", "awAqi", "What the AQI represents", "Multiple response: what parents think the AQI measures", SD(f_multi("beleifs_air_pollute", "belief_air", sort=False)), var="beleifs_air_pollute", opt={"color": 1}),
            card("bar", "awRate", "Outdoor air quality rating", "0 = not polluted at all, 10 = extremely polluted", SD(lambda d: {"items": items([str(i) for i in range(11)], [int((num(d, "outdoor_air_quality") == i).sum()) for i in range(11)], int(num(d, "outdoor_air_quality").notna().sum())), "base": int(num(d, "outdoor_air_quality").notna().sum())} if num(d, "outdoor_air_quality").notna().any() else None), var="outdoor_air_quality", opt={"color": 2}),
            card("donut", "awCheck", "How often parents check air quality", "Website or app monitoring of outdoor pollution", SD(f_dist("check_air_quality", "check_air", hide=(98,))), var="check_air_quality", opt={"colors": [6, 3, 1, 4, 2, "muted"]})]},
        {"t": "grid", "cols": 2, "cards": [
            card("hbar", "awSource", "Primary source of AQI information", "Where parents get air quality information", SD(f_dist("primary_source", "source", drop_zero=True)), var="primary_source", opt={"color": 3}),
            card("hbar", "awIndoor", "Indoor vs outdoor pollution", "How polluted the home air is believed to be, compared with outdoors", SD(f_dist("pollute_believe", "belief", hide=(98,))), var="pollute_believe", opt={"color": 5})]},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "awRisk", "Perceived increase in child's respiratory risk", "Effect of indoor and school air pollution on acute respiratory illness", SD(f_dist("risk_air_child", "disease_risk", hide=(98,))), var="risk_air_child", opt={"color": 2}),
            card("likert", "awExam", "Expected exam gain: air conditioner vs air purifier", "How much higher parents think their child would score with each in the classroom all year",
                 SD(f_likert([("Air conditioner", "score_exams_ac"), ("Air purifier", "score_exams_ap")], "score")), var="score_exams_ac / score_exams_ap", opt={"pal": "seq"})]},
        {"t": "grid", "cols": 2, "cards": [
            card("hbar", "awKnow", "Measures parents know about", "Multiple response, unprompted: ways to reduce health harm from air pollution", SD(f_multi("measure_air_pollution", "measures", top=10)), var="measure_air_pollution", opt={"color": 6}),
            card("hbar", "awDo", "Precautions taken for the child", "Multiple response, unprompted: what parents ask their child to do", SD(f_multi("actions_protect_c", "measures", top=10)), var="actions_protect_c", opt={"color": 4})]},
        {"t": "grid", "cols": 3, "cards": [
            card("donut", "awSeason", "Would a November test score be lower?", "Same preparation, same child: April test vs November test", SD(f_dist("expect_score", "yesno_sure")), var="expect_score", opt={"colors": [6, 2, "muted"]}),
            card("donut", "awSchool", "School has air quality measures", "Fans, air conditioners, open-window policy or similar", SD(f_dist("air_school", "yesno_sure2")), var="air_school", opt={"colors": [6, 2, "muted"]}),
            card("bar", "awRank", "Priority of school air quality", "Compared with better teachers, facilities or books", SD(f_dist("rank_improve", "rank")), var="rank_improve", opt={"color": 5})]},
        {"t": "insights", "d": SD(aware_ins), "row": True},
        {"t": "note", "html": "“Don't know” and “Refused” responses are shown but excluded from the summary indicators in the strip above. Multiple-response items report the share of respondents selecting each option, so bars can add to more than 100%."},
    ]})

# ====================================================================== PANEL 6 - WILLINGNESS TO PAY
BIN_E = [0, 1, 251, 501, 751, 1001, 1501, 2000, 2001]
BIN_L = ["Rs 0", "1–250", "251–500", "501–750", "751–1,000", "1,001–1,500", "1,501–1,999", "Rs 2,000"]
BIN_E[0] = -1


def wtp_kpis(df):
    b = df["_bid"]
    tp = thr_price(b)
    return [
        K("💰", fmt_rs(avg(b)), "Mean Willingness to Pay", f"Median {fmt_rs(med(b))}", "up", "green"),
        K("🙋", fmt_p(sh(b, lambda s: s > 0)), "Bid Above Zero", f"{fmt_p(sh(b, lambda s: s >= 2000))} bid the Rs 2,000 maximum", "neutral", "teal"),
        K("🎯", fmt_rs(tp), "Price Clearing the 30% Threshold", "Highest price at which 30% still contribute", "up", "purple"),
        K("🔄", fmt_p(sh(num(df, "change_pay"), lambda s: s == 1)), "Revised Their Bid", f"Mean bid {fmt_rs(avg(df['_bid0']))} → {fmt_rs(avg(b))}", "neutral", "amber"),
        K("🎲", fmt_p(pct(int(df["_hit"].sum()), int(df["_hit"].notna().sum())) if df["_hit"].notna().any() else None), "Would Have Contributed", "Bid at or above the random draw", "navy", "navy"),
        K("💸", fmt_rs(avg(df["_paid"])), "Mean Amount Paid", "Realised by the random-price rule", "navy", ""),
    ]


def others_belief(df):
    tot = pd.Series(np.nan, index=df.index)
    for c in ("total_parents", "total_parents_2"):
        tot = tot.fillna(num(df, c))
    rem = tot - 1
    parts = [("Not surveyed", "other_parents_fail"), ("Rs 0", "other_parents_0"), ("Rs 1–500", "other_parents_500"), ("Rs 501–1,000", "other_parents_1000"),
             ("Rs 1,001–1,500", "other_parents_1500"), ("Rs 1,501–1,999", "other_parents_1999"), ("Rs 2,000", "other_parents_2000")]
    its = []
    for lab, c in parts:
        v = num(df, c).fillna(num(df, c + "_2")) / rem
        v = v[rem > 0].dropna()
        if len(v):
            its.append({"l": lab, "v": r1(100 * float(v.mean())), "n": len(v)})
    return {"items": its, "base": int(rem.notna().sum())} if its else None


def wtp_ins(df):
    b = df["_bid"]
    exp_ = None
    tot = pd.Series(np.nan, index=df.index)
    for c in ("total_parents", "total_parents_2"):
        tot = tot.fillna(num(df, c))
    op = num(df, "other_parents").fillna(num(df, "other_parents_2"))
    ratio = (op / (tot - 1)).replace([np.inf, -np.inf], np.nan).dropna()
    if len(ratio):
        exp_ = 100 * float(ratio.mean())
    real = pct(int(df["_hit"].sum()), int(df["_hit"].notna().sum())) if df["_hit"].notna().any() else None
    out = [{"cls": "primary", "html": f"<strong>The threshold:</strong> a classroom gets a purifier only if 30% of parents contribute. Bids clear that bar up to <strong>{fmt_rs(thr_price(b))}</strong>; above it the classroom would not reach 30%."}]
    if exp_ is not None and real is not None:
        out.append({"cls": "navy", "html": f"<strong>Beliefs about other parents:</strong> parents expect <strong>{exp_:.0f}%</strong> of other parents to contribute; on the random-price rule <strong>{real:.0f}%</strong> of respondents' own bids would contribute. {'Parents underestimate their peers.' if exp_ < real else 'Parents overestimate their peers.'}"})
    return out


# WTP by respondent group (not affected by the segment selector)
def wtp_groups():
    rows = []
    add = lambda grp, lab, mask, slot: rows.append((grp, lab, C.loc[mask, "_bid"], slot))
    for v, l in CH.get("area", []):
        add("Area", l, num(C, "area_class") == v, 4)
    for v, l in CH.get("gender", []):
        if v in (1, 2):
            add("Gender", l, num(C, "gender") == v, 7)
    q = C["_income"].quantile([1 / 3, 2 / 3]).values if C["_income"].notna().sum() >= 6 else None
    if q is not None:
        add("Household income", "Lowest third", C["_income"] <= q[0], 3)
        add("Household income", "Middle third", (C["_income"] > q[0]) & (C["_income"] <= q[1]), 3)
        add("Household income", "Highest third", C["_income"] > q[1], 3)
    add("Owns an air conditioner", "Yes", num(C, "asset_ac") == 1, 5); add("Owns an air conditioner", "No", num(C, "asset_ac") == 2, 5)
    add("Respiratory illness at home", "Yes", num(C, "hh_resp_ill") == 1, 2); add("Respiratory illness at home", "No", num(C, "hh_resp_ill") == 2, 2)
    add("Perceived child risk", "High (≥50% higher)", num(C, "risk_air_child").isin([4, 5]), 8); add("Perceived child risk", "Low or none", num(C, "risk_air_child").isin([1, 2, 3]), 8)
    for v in (0, 1, 2):
        add("Study arm", ARM_LAB[v], C["_arm"] == v, [7, 3, 5][v])
    out = []
    for g, l, s, slot in rows:
        m = mean_ci(s)
        if m and m["n"] >= MIN_CELL:
            out.append({"g": g, "l": l, "m": r1(m["m"], 0), "lo": r1(m["lo"], 0) if m["lo"] is not None else None,
                        "hi": r1(m["hi"], 0) if m["hi"] is not None else None, "n": m["n"], "slot": slot})
    return {"rows": out, "unit": "Rs", "overall": r1(avg(C["_bid"]), 0)} if out else None


def d_demand(df):
    y = demand(df["_bid"])
    return {"x": [f"Rs {p:,}" for p in PRICES], "base": int(df["_bid"].notna().sum()),
            "series": [{"name": "Share willing to contribute at least this much", "slot": 2, "y": y, "fill": True, "pts": True},
                       {"name": "30% needed for the purifier", "slot": 1, "y": [30] * len(PRICES), "dash": True}]} if y else None


PANELS.append({
    "id": "wtp", "tab": "💰 Willingness to Pay", "eyebrow": "Section 06 · Incentive-compatible elicitation (BDM)", "title": "Willingness to Pay for a Classroom Air Purifier",
    "blurb": "Each parent states the most they would contribute towards an air purifier for their child's classroom, then a random price is drawn: they pay the random price only if their bid is at least as high, which makes an honest bid the best strategy. The classroom gets a purifier if 30% of parents contribute.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(wtp_kpis)},
        {"t": "grid", "cols": 1, "cards": [
            card("bar", "wtDist", "Distribution of final bids", "Amount each parent was willing to contribute after any revision, in rupees", SD(f_hist(lambda d: d["_bid"], BIN_E, BIN_L)), var="contribute_will / change_pay_w", opt={"color": 6, "fmt": "pct"}, full=True)]},
        {"t": "grid", "cols": 2, "cards": [
            card("lines", "wtDemand", "Demand curve", "Share of parents willing to contribute at least each price; the dashed line is the 30% needed for a purifier", SD(d_demand), opt={"yfmt": "pct", "max": 100, "xtitle": "Contribution price", "ytitle": "% willing to contribute"}, tall=True),
            card("hbar", "wtGroup", "Mean bid by respondent group", "Mean final bid with 95% confidence interval; not affected by the segment selector", wtp_groups(), opt={"grouped": True}, tall=True, auto=True)]},
        {"t": "grid", "cols": 3, "cards": [
            card("bar", "wtOthers", "What parents expect other parents to do", "Average expected split of the other parents in the class, % of class", SD(others_belief), opt={"color": 5, "fmt": "num1", "sfx": "%"}),
            card("donut", "wtChange", "Revised their bid after seeing the result", "Would you like to change how much you are willing to contribute?", SD(f_dist("change_pay", "yesno")), var="change_pay", opt={"colors": [4, "muted"]}),
            card("donut", "wtCertain", "Certainty about the bid", "How certain parents are of their choice", SD(f_dist("certain_choice", "choice", drop_zero=True)), var="certain_choice", opt={"colors": [6, 3, 1, 4, "muted"]})]},
        {"t": "grid", "cols": 2, "cards": [
            card("hbar", "wtWhy", "Why not pay more", "Reason given by parents who bid below the Rs 2,000 maximum and kept their bid", SD(f_dist("more_pay", "reason", drop_zero=True)), var="more_pay", opt={"color": 2}),
            card("bar", "wtPractice", "Practice round bids", "Rs 0–20 practice round before the real exercise", SD(f_hist(lambda d: num(d, "contribute_will_practice"), [-1, 1, 6, 11, 16, 21], ["Rs 0", "1–5", "6–10", "11–15", "16–20"])), var="contribute_will_practice", opt={"color": 1})]},
        {"t": "grid", "cols": 2, "cards": [
            card("donut", "wtSel", "If selected and the price is above the bid", "Comprehension: is the Rs 1,000 bonus still received?", SD(f_dist("selected_parent", "yesno")), var="selected_parent", opt={"colors": [6, 2]}),
            card("donut", "wtNot", "If not selected: which bid is used", "Comprehension of the second-stage rule", SD(f_dist("not_selected_parent", "bid_use")), var="not_selected_parent", opt={"colors": [6, 2, 4]})]},
        {"t": "insights", "d": SD(wtp_ins), "row": True},
        {"t": "note", "html": "Final bid = the revised bid where a parent changed it, otherwise the original. “Would have contributed” compares the bid with the random price drawn in the interview. The 30% threshold price is the highest price at which at least 30% of the segment would still contribute."},
    ]})

# ====================================================================== PANEL 7 - TREATMENT EFFECTS
arm_masks = {v: C["_arm"] == v for v in (0, 1, 2)}
ARM_SLOT = {0: 7, 1: 3, 2: 5}
n_arm = {v: int(arm_masks[v].sum()) for v in (0, 1, 2)}
te_kpis = []
for v in (0, 1, 2):
    m = avg(C.loc[arm_masks[v], "_bid"])
    dlt = None if v == 0 or avg(C.loc[arm_masks[0], "_bid"]) is None or m is None else m - avg(C.loc[arm_masks[0], "_bid"])
    te_kpis.append(K(["⚪", "🎬", "🎞️"][v], fmt_rs(m), f"{ARM_LAB[v]}: Mean Bid", f"n = {n_arm[v]}" + ("" if dlt is None else f" · {dlt:+,.0f} vs Control"), "up" if (dlt or 0) > 0 else "neutral", ["navy", "teal", "purple"][v]))
ords = {v: C["_ord"] == v for v in (1, 2)}
_o = welch(C.loc[ords[1], "_bid"], C.loc[ords[2], "_bid"])
te_kpis.append(K("🔀", fmt_rs(avg(C.loc[ords[2], "_bid"])), "Beliefs After Bid: Mean Bid", "" if _o is None else f"{_o['diff']:+,.0f} vs beliefs-before (p = {_o['p']:.2f})", "neutral", "amber"))

OUTCOMES = [
    ("Mean bid (Rs)", "rs", lambda d: d["_bid"]),
    ("Mean bonus-round bid (Rs)", "rs", lambda d: d["_bonus"]),
    ("Bid above zero", "pp", lambda d: (d["_bid"] > 0).astype(float).where(d["_bid"].notna())),
    ("Bid at the Rs 2,000 maximum", "pp", lambda d: (d["_bid"] >= 2000).astype(float).where(d["_bid"].notna())),
    ("Sees ≥50% higher child risk", "pp", lambda d: num(d, "risk_air_child").where(num(d, "risk_air_child").isin([1, 2, 3, 4, 5])).map(lambda v: np.nan if pd.isna(v) else float(v >= 4))),
    ("Expects ≥10% exam gain from purifier", "pp", lambda d: num(d, "score_exams_ap").map(lambda v: np.nan if pd.isna(v) else float(v >= 4))),
    ("Supports publicly funded purifiers", "pp", lambda d: num(d, "support_policy").map(lambda v: np.nan if pd.isna(v) else float(v in (1, 2)))),
    ("Supports a small mandatory school fee", "pp", lambda d: num(d, "small_fee").map(lambda v: np.nan if pd.isna(v) else float(v == 1))),
    ("Air pollution should be a top school priority", "pp", lambda d: num(d, "top_prio").map(lambda v: np.nan if pd.isna(v) else float(v == 1))),
]
CMPS = [("Video 1 vs Control", lambda d: d["_arm"] == 0, lambda d: d["_arm"] == 1),
        ("Video 2 vs Control", lambda d: d["_arm"] == 0, lambda d: d["_arm"] == 2),
        ("Video 2 vs Video 1", lambda d: d["_arm"] == 1, lambda d: d["_arm"] == 2),
        ("Any video vs Control", lambda d: d["_arm"] == 0, lambda d: d["_arm"].isin([1, 2]))]
eff_rows = []
for cname, fa, fb in CMPS:
    for oname, unit, getter in OUTCOMES:
        s = getter(C)
        w = welch(s[fa(C)], s[fb(C)])
        if not w:
            continue
        k_ = 1 if unit == "rs" else 100
        fmt_v = (lambda x: f"{x:,.0f}") if unit == "rs" else (lambda x: f"{x:.0f}%")
        fmt_d = (lambda x: f"{x:+,.0f}") if unit == "rs" else (lambda x: f"{x:+.1f} pp")
        eff_rows.append({"cmp": cname, "o": oname, "a": {"v": r1(s[fa(C)].mean() * k_), "t": fmt_v(s[fa(C)].mean() * k_)},
                         "b": {"v": r1(s[fb(C)].mean() * k_), "t": fmt_v(s[fb(C)].mean() * k_)},
                         "diff": {"v": r1(w["diff"] * k_), "t": fmt_d(w["diff"] * k_)},
                         "ci": {"v": r1(w["lo"] * k_), "t": f"{fmt_d(w['lo'] * k_)} to {fmt_d(w['hi'] * k_)}"},
                         "p": {"v": round(w["p"], 3), "t": f"{w['p']:.2f}" if w["p"] >= 0.01 else "<0.01", "sig": bool(w["p"] < 0.05)},
                         "n": f"{w['na']} / {w['nb']}"})

# balance table
BAL = [("Age (mean, years)", "num", lambda d: d["_age"]),
       ("Monthly income (median, Rs)", "med", lambda d: d["_income"]),
       ("Household size (mean)", "num", lambda d: d["_hhsize"]),
       ("Female respondent", "sh", lambda d: (num(d, "gender") == 2).astype(float).where(num(d, "gender").notna())),
       ("Urban area", "sh", lambda d: (num(d, "area_class") == 2).astype(float).where(num(d, "area_class").notna())),
       ("Owns an air conditioner", "sh", lambda d: (num(d, "asset_ac") == 1).astype(float).where(num(d, "asset_ac").notna())),
       ("Respiratory illness at home", "sh", lambda d: (num(d, "hh_resp_ill") == 1).astype(float).where(num(d, "hh_resp_ill").notna())),
       ("Outdoor pollution rating (mean, 0–10)", "num", lambda d: num(d, "outdoor_air_quality")),
       ("Child's grade (mean)", "num", lambda d: d["_grade"])]
bal_rows = []
for name, kind, getter in BAL:
    s = getter(C)
    groups = [s[arm_masks[v]] for v in (0, 1, 2)]
    if kind == "sh":
        vals = [sh(g, lambda x: x == 1) for g in groups]; p_ = chi2_p(groups); t = lambda x: "—" if x is None else f"{x:.0f}%"
    elif kind == "med":
        vals = [med(g) for g in groups]; p_ = anova_p([np.log(g.where(g > 0)) for g in groups]); t = lambda x: "—" if x is None else f"{x:,.0f}"
    else:
        vals = [avg(g) for g in groups]; p_ = anova_p(groups); t = lambda x: "—" if x is None else f"{x:.1f}"
    bal_rows.append({"o": name, "c": {"v": vals[0], "t": t(vals[0])}, "v1": {"v": vals[1], "t": t(vals[1])}, "v2": {"v": vals[2], "t": t(vals[2])},
                     "p": {"v": None if p_ is None else round(p_, 3), "t": "—" if p_ is None else f"{p_:.2f}", "flag": bool(p_ is not None and p_ < 0.10)}})
n_imb = sum(1 for r in bal_rows if r["p"]["flag"])

ci_arm = {"rows": [], "unit": "Rs"}
for v in (0, 1, 2):
    m = mean_ci(C.loc[arm_masks[v], "_bid"])
    if m:
        ci_arm["rows"].append({"g": "Study arm", "l": ARM_LAB[v], "m": r1(m["m"], 0), "lo": r1(m["lo"], 0) if m["lo"] is not None else None, "hi": r1(m["hi"], 0) if m["hi"] is not None else None, "n": m["n"], "slot": ARM_SLOT[v]})
ci_ord = {"rows": [], "unit": "Rs"}
for arm in (0, 1, 2):
    for o in (1, 2):
        m = mean_ci(C.loc[arm_masks[arm] & ords[o], "_bid"])
        if m and m["n"] >= 3:
            ci_ord["rows"].append({"g": ARM_LAB[arm], "l": "Asked before own bid" if o == 1 else "Asked after own bid", "m": r1(m["m"], 0), "lo": r1(m["lo"], 0) if m["lo"] is not None else None, "hi": r1(m["hi"], 0) if m["hi"] is not None else None, "n": m["n"], "slot": ARM_SLOT[arm]})
dem_arm = {"x": [f"Rs {p:,}" for p in PRICES], "base": N_C, "series": []}
for v in (0, 1, 2):
    y = demand(C.loc[arm_masks[v], "_bid"])
    if y:
        dem_arm["series"].append({"name": f"{ARM_LAB[v]} (n = {n_arm[v]})", "slot": ARM_SLOT[v], "y": y, "pts": True})
dem_arm["series"].append({"name": "30% needed for the purifier", "slot": 1, "y": [30] * len(PRICES), "dash": True, "thin": True})
thr_arm = {v: thr_price(C.loc[arm_masks[v], "_bid"]) for v in (0, 1, 2)}
if len(dem_arm["series"]) < 2:
    dem_arm = None
ci_arm = ci_arm if ci_arm["rows"] else None
ci_ord = ci_ord if ci_ord["rows"] else None
bal_data = {"cols": [{"k": "o", "l": "Characteristic"}, {"k": "c", "l": "Control", "num": True}, {"k": "v1", "l": "Video 1", "num": True}, {"k": "v2", "l": "Video 2", "num": True}, {"k": "p", "l": "p-value", "num": True, "pflag": True}], "rows": bal_rows} if N_C >= 3 else None
eff_data = {"cols": [{"k": "o", "l": "Outcome"}, {"k": "a", "l": "Comparison group", "num": True}, {"k": "b", "l": "Treatment group", "num": True}, {"k": "diff", "l": "Difference", "num": True}, {"k": "ci", "l": "95% CI", "num": True}, {"k": "p", "l": "p-value", "num": True, "psig": True}, {"k": "n", "l": "n (comparison / treatment)", "num": True}], "rows": eff_rows} if eff_rows else None

te_ins = [
    {"cls": "navy", "html": f"<strong>Randomisation:</strong> " + (f"<strong>{n_imb}</strong> of {len(bal_rows)} background characteristics differ across arms at the 10% level: expected by chance in roughly one in ten, but worth a look if it grows." if n_imb else f"none of the {len(bal_rows)} background characteristics differ across arms at the 10% level, so the three groups look comparable.")},
    {"cls": "teal", "html": "<strong>Price clearing 30%:</strong> " + " · ".join(f"{ARM_LAB[v]} <strong>{fmt_rs(thr_arm[v])}</strong>" for v in (0, 1, 2)) + ". A higher price means more classrooms reach the threshold at that contribution level."},
]
PANELS.append({
    "id": "effects", "tab": "🧪 Treatment Effects", "eyebrow": "Section 07 · Randomised information experiment", "title": "Treatment Effects & Randomisation",
    "blurb": "Households were randomly assigned to no video (Control), Video 1 or Video 2 before the willingness-to-pay exercise, and to hearing about other parents' behaviour before or after their own bid. Differences between arms are the field read of what the videos change.",
    "seg": False, "blocks": [
        {"t": "insight_static", "cls": "navy", "html": "<strong>🧪 Design:</strong> <strong>Control</strong> sees no video. <strong>Video 1</strong> and <strong>Video 2</strong> are two short films about classroom air purifiers shown before the bid. Assignment is pre-drawn in the sampling frame with a fixed seed, so arm is unrelated to who the enumerator visits."},
        {"t": "kpis", "d": te_kpis},
        {"t": "grid", "cols": 2, "cards": [
            card("lines", "teDemand", "Demand curve by study arm", "Share of parents willing to contribute at least each price, by arm; the dashed line is the 30% needed", dem_arm, opt={"yfmt": "pct", "max": 100, "xtitle": "Contribution price", "ytitle": "% willing to contribute"}, tall=True),
            card("hbar", "teArmCi", "Mean bid by study arm", "Mean final bid with 95% confidence interval", ci_arm, opt={"grouped": True}, tall=True)]},
        {"t": "grid", "cols": 1, "cards": [
            card("table", "teEff", "Estimated differences between arms", "Difference in means (Welch t-test). Rupees for bids; percentage points for shares. Shaded p-values are below 0.05. With small samples treat these as indicative.",
                 eff_data, full=True, opt={"filterKey": "cmp", "filterVals": [c[0] for c in CMPS], "noSort": True})]},
        {"t": "grid", "cols": 2, "cards": [
            card("table", "teBal", "Randomisation balance", "Background characteristics by arm; p-value from an F-test (numeric) or chi-squared test (shares). p under 0.10 is flagged to verify.",
                 bal_data, opt={"noSort": True}),
            card("hbar", "teOrder", "Order effect: beliefs before or after the bid", "Mean bid with 95% CI when parents were asked about other parents before vs after their own bid", ci_ord, opt={"grouped": True}, tall=True, auto=True)]},
        {"t": "insights", "d": te_ins, "row": True},
        {"t": "note", "html": "Estimates use completed interviews only, counted once per household. Confidence intervals and p-values are unadjusted for multiple comparisons; with about 33 households per arm they are wide, and they narrow as fieldwork completes."},
    ]})

# ====================================================================== PANEL 8 - FAIRNESS & COLLECTIVE ACTION
PRIORITIES = [("Education", "edu"), ("Healthcare", "heal"), ("Air quality", "airy"), ("Job creation", "jobs"), ("Elderly care", "elder_care")]


def prio(df):
    its = []
    for l, c in PRIORITIES:
        s = num(df, c).dropna()
        if len(s):
            its.append({"l": l, "v": r1(float((s <= 2).mean() * 100)), "n": int((s <= 2).sum()), "p": r1(float((s <= 2).mean() * 100)), "m": r1(float(s.mean()), 2)})
    its.sort(key=lambda i: -i["v"])
    return {"items": its, "base": int(num(df, "edu").notna().sum())} if its else None


def fair_kpis(df):
    return [
        K("⚖️", fmt_p(top2(num(df, "pay_extra"), [1, 2])), "“It's the School's Job to Pay”", "Agree or strongly agree", "neutral", "navy"),
        K("🫁", fmt_p(top2(num(df, "right_to_pay"), [1, 2])), "Clean Air Is a Right", "Agree or strongly agree", "neutral", "teal"),
        K("🚶", fmt_p(sh(num(df, "others_will"), lambda s: s == 1)), "Would Still Contribute if Others Did", "Even if the purifier is already guaranteed", "neutral", "purple"),
        K("🏛️", fmt_p(pct(int(num(df, "support_policy").isin([1, 2]).sum()), int(num(df, "support_policy").notna().sum())) if num(df, "support_policy").notna().any() else None), "Support Publicly Funded Purifiers", "Yes, regardless of cost or if small", "up", "green"),
        K("🎯", fmt_p(sh(num(df, "top_prio"), lambda s: s == 1)), "Air Should Be a Top School Priority", "Yes", "up", "amber"),
    ]


def fair_ins(df):
    pe = top2(num(df, "pay_extra"), [1, 2]); eq = top2(num(df, "equally_pay"), [1, 2])
    ow = sh(num(df, "others_will"), lambda s: s == 1)
    return [{"cls": "amber", "html": f"<strong>Who should pay:</strong> <strong>{fmt_p(pe)}</strong> of parents say the school, not parents, should fund purifiers, yet <strong>{fmt_p(eq)}</strong> would prefer everyone contribute equally to a shared scheme."},
            {"cls": "purple", "html": f"<strong>Free-riding:</strong> if enough others were already paying, <strong>{fmt_p(ow)}</strong> would still contribute, which suggests a genuine willingness to pay beyond the strategic incentive to wait."}]


PANELS.append({
    "id": "fair", "tab": "⚖️ Fairness & Action", "eyebrow": "Section 08 · Public goods & collective action", "title": "Fairness, Free-riding & Collective Action",
    "blurb": "Whether parents think families or schools should pay for clean classroom air, whether they would free-ride if others covered the cost, how air quality ranks against other government priorities, and their support for policy options.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(fair_kpis)},
        {"t": "grid", "cols": 1, "cards": [
            card("likert", "faAgree", "Who should pay for classroom air purifiers", "Level of agreement with each statement", SD(f_likert([("Parents should not be asked to pay extra: it's the school's job", "pay_extra"), ("Clean classroom air is a right, not something families should pay for", "right_to_pay"), ("Every parent should contribute equally rather than rely on voluntary support", "equally_pay")], "agree")), var="pay_extra · right_to_pay · equally_pay", opt={"pal": "div"}, full=True)]},
        {"t": "grid", "cols": 3, "cards": [
            card("hbar", "faFree", "Free-rider questions", "Share answering yes", SD(f_yes([("Would contribute if enough others already had", "others_will"), ("Would still contribute to maintenance if provided", "will_air")])), var="others_will · will_air", opt={"color": 5, "fmt": "pct"}),
            card("bar", "faKnow", "Other parents known personally", "Parents of children in the same classroom the respondent knows", SD(f_hist(lambda d: num(d, "other_parent"), [0, 1, 4, 8, 13, 1000], ["0", "1–3", "4–7", "8–12", "13+"])), var="other_parent", opt={"color": 1}),
            card("hbar", "faPrio", "Government priorities: ranked first or second", "Share ranking each area in their top two (1 = most priority)", SD(prio), var="edu · heal · airy · jobs · elder_care", opt={"color": 2, "fmt": "pct"})]},
        {"t": "grid", "cols": 3, "cards": [
            card("donut", "faNgo", "Preferred NGO for a Rs 3,000 donation", "Hypothetical choice between three organisations or none", SD(f_dist("hypo_q", "ngo")), var="hypo_q", opt={"colors": [1, 2, 3, "muted"]}),
            card("donut", "faPol", "Support for publicly funded purifiers", "Installing purifiers in all primary classrooms with partial funding from taxes", SD(f_dist("support_policy", "support")), var="support_policy", opt={"colors": [6, 3, 4, "muted"]}),
            card("hbar", "faSchool", "School priorities", "Share answering yes", SD(f_yes([("Support a mandatory small tuition fee", "small_fee"), ("Pay more if it improved learning", "improve_learn"), ("Air pollution should be a top school priority", "top_prio")])), opt={"color": 6, "fmt": "pct"})]},
        {"t": "insights", "d": SD(fair_ins), "row": True},
        {"t": "note", "html": "Agreement scales run from strongly agree to strongly disagree and are shown as shares of valid responses. Priority ranks run from 1 (most priority) to 5; the chart shows the share placing each area first or second."},
    ]})

# ====================================================================== PANEL 9 - TIME PREFERENCES & PAYOUT
def time_kpis(df):
    disc = df["_disc"]
    return [
        K("⏳", f"{med(disc):.0f}%" if med(disc) is not None else "—", "Median Implied Annual Discount Rate", "Reward demanded to wait one year", "neutral", "navy"),
        K("🐢", r1(avg(num(df, "tp_qual"))) or "—", "Patience Score (0–10)", "10 = very willing to give up something today", "up", "teal"),
        K("🛑", fmt_p(sh(num(df, "tp1"), lambda s: s == 1)), "Take Rs 2,000 Today over ~Rs 4,100", "First question, one year later", "neutral", "amber"),
        K("✔️", fmt_p(sh(num(df, "tp_check"), lambda s: s == 1)), "Consistent Check Answers", "Prefer 2,000 today to 2,000 in a year", "up", "green"),
        K("📲", fmt_rs(avg(num(df, "topup_amount"))), "Mean Mobile Top-up", "Rs 2,000 less any contribution paid", "navy", "purple"),
    ]


def time_ins(df):
    return [{"cls": "navy", "html": f"<strong>Impatience:</strong> the typical parent demands about <strong>{f'{med(df['_disc']):.0f}%' if med(df['_disc']) is not None else '—'}</strong> a year to give up money today, a very high implicit rate that is common in low-income settings and matters for how much weight future health benefits receive."},
            {"cls": "teal", "html": f"<strong>Data quality:</strong> <strong>{fmt_p(sh(num(df, 'tp_check'), lambda s: s == 1))}</strong> pass the consistency check (2,000 today vs the same amount in a year); the remainder may not have understood the question."}]


PANELS.append({
    "id": "time", "tab": "⏳ Time & Payout", "eyebrow": "Section 09 · Time preferences & incentive payout", "title": "Time Preferences & Mobile Top-up",
    "blurb": "Three adaptive choices between Rs 2,000 today and a larger amount in one year place each respondent's patience on a common scale. Participants also receive a mobile top-up worth Rs 2,000 less anything they contributed.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(time_kpis)},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "tpDisc", "Implied annual discount rate", "The reward, per year, at which each respondent switches from “today” to “in one year”", SD(f_hist(lambda d: d["_disc"], [-1000, 25, 50, 100, 200, 1e9], ["<25%", "25–49%", "50–99%", "100–199%", "200%+"])), var="tp_bin", opt={"color": 4, "fmt": "pct"}),
            card("bar", "tpQual", "Willingness to give up something today", "0 = completely unwilling, 10 = very willing", SD(lambda d: {"items": items([str(i) for i in range(11)], [int((num(d, "tp_qual") == i).sum()) for i in range(11)], int(num(d, "tp_qual").notna().sum())), "base": int(num(d, "tp_qual").notna().sum())} if num(d, "tp_qual").notna().any() else None), var="tp_qual", opt={"color": 6})]},
        {"t": "grid", "cols": 3, "cards": [
            card("donut", "tpFirst", "First choice: Rs 2,000 today or more in a year", "The opening question of the bisection", SD(f_dist("tp1", "timeopt", hide=(98,))), var="tp1", opt={"colors": [2, 3, "muted"]}),
            card("donut", "tpCheck", "Consistency check", "Rs 2,000 today or the same Rs 2,000 in one year", SD(f_dist("tp_check", "timeopt", hide=(98,))), var="tp_check", opt={"colors": [6, 2, "muted"]}),
            card("donut", "tpNet", "Mobile network for top-up", "Network the respondent uses", SD(f_dist("mobile_network", "network_pk", drop_zero=True)), var="mobile_network", opt={"colors": [1, 2, 3, 4, "muted"]})]},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "tpTop", "Top-up amount", "Rs 2,000 less the contribution paid under the random-price rule", SD(f_hist(lambda d: num(d, "topup_amount"), [0, 1000, 1500, 1800, 2000, 2001], ["<1,000", "1,000–1,499", "1,500–1,799", "1,800–1,999", "2,000"])), var="topup_amount", opt={"color": 5, "fmt": "pct"}),
            card("bar", "tpPaid", "Contribution actually paid", "Under the random-price rule; zero if the price drawn was above the bid", SD(f_hist(lambda d: d["_paid"], [-1, 1, 251, 501, 1001, 2001], ["Rs 0", "1–250", "251–500", "501–1,000", "1,001+"])), var="you_paid_amount_main", opt={"color": 3, "fmt": "pct"})]},
        {"t": "insights", "d": SD(time_ins), "row": True},
        {"t": "note", "html": "The implied discount rate is read from the final node of the three-step bisection (Rs 2,100 to Rs 8,000 offered in one year). Responses marked invalid or refused are excluded. Mobile numbers are collected for the top-up only and are never loaded into this dashboard."},
    ]})

# ====================================================================== PANEL 10 - HOUSEHOLD TRACKER
ST_LAB = {1: "Completed", 2: "Partial", 3: "Refused", 4: "Refused", 5: "Refused", 6: "Refused", 7: "Not reached", 8: "Not reached", 9: "Ineligible", 10: "Not reached"}
tr_rows = []
ids = list(frame.index) if frame is not None else sorted(D["_hh"].dropna().astype(int).unique())
Ds = D.sort_values("_dt")
for hid in ids:
    g = Ds[Ds["_hh"] == hid]
    if len(g):
        st = ST_LAB.get(int(g["_status"].iloc[-1]) if pd.notna(g["_status"].iloc[-1]) else 0, "Other")
        if (g["_status"] == 1).any():
            st = "Completed"
        last = g.iloc[-1]
        tr_rows.append({"id": int(hid), "arm": int(frame.loc[hid, "ap_lm_arm"]) if frame is not None else (int(last["_arm"]) if pd.notna(last["_arm"]) else -1),
                        "o": int(frame.loc[hid, "order_key"]) if frame is not None else (int(last["_ord"]) if pd.notna(last["_ord"]) else 0),
                        "st": st, "att": int(len(g)), "e": last["_enum_lab"], "last": fdate(last["_dt"]) if pd.notna(last["_dt"]) else "—",
                        "min": r1(med(g.loc[g["_status"] == 1, "_min"]) if (g["_status"] == 1).any() else None, 0)})
    else:
        tr_rows.append({"id": int(hid), "arm": int(frame.loc[hid, "ap_lm_arm"]) if frame is not None else -1,
                        "o": int(frame.loc[hid, "order_key"]) if frame is not None else 0, "st": "Not visited", "att": 0, "e": "—", "last": "—", "min": None})
st_counts = {}
for r in tr_rows:
    st_counts[r["st"]] = st_counts.get(r["st"], 0) + 1
tr_kpis = [K("🏠", f"{len(tr_rows):,}", "Households in Frame", f"{ARM_LAB[0]} {arm_target[0]} · {ARM_LAB[1]} {arm_target[1]} · {ARM_LAB[2]} {arm_target[2]}", "navy", "navy"),
           K("✅", st_counts.get("Completed", 0), "Completed", f"{pct(st_counts.get('Completed', 0), len(tr_rows), 0):.0f}% of frame", "up", "green"),
           K("🔁", st_counts.get("Not reached", 0) + st_counts.get("Partial", 0), "To Revisit", "Not reached or partial", "neutral", "amber"),
           K("🚫", st_counts.get("Refused", 0) + st_counts.get("Ineligible", 0), "Refused or Ineligible", "No further visit planned", "down", ""),
           K("🕓", st_counts.get("Not visited", 0), "Not Yet Visited", "Waiting for a first attempt", "neutral", "teal")]
PANELS.append({
    "id": "tracker", "tab": "🏠 Household Tracker", "eyebrow": "Section 10 · Operational tracker", "title": "Household Completion Tracker",
    "blurb": "Status of every household in the sampling frame: which are complete, which need a revisit and which have not been visited yet. Search by household ID, filter by status, arm or enumerator, click a column to sort, or export the current view to CSV.",
    "seg": False, "blocks": [
        {"t": "kpis", "d": tr_kpis},
        {"t": "blocks", "head": "📐 Progress by assignment block", "d": blocks, "foot": "Each card is one study arm and question-order combination from the sampling frame. Target = households assigned to the block; completed = interviews done. This is the quickest check that all six blocks fill evenly."},
        {"t": "tracker", "d": {"rows": tr_rows, "enums": sorted({r["e"] for r in tr_rows if r["e"] != "—"}), "arms": {str(k): v for k, v in ARM_LAB.items()}}},
        {"t": "note", "html": "One row per household ID in the frame. Status reflects the latest attempt, or Completed if any attempt was completed. Only the household ID, assigned arm and visit outcome are shown: no names, addresses or phone numbers."},
    ]})

# ====================================================================== payload, PII scan, encrypt
DATA = {"meta": META, "segments": SEG_META, "panels": PANELS}


def clean(o):
    if isinstance(o, dict):
        return {str(k): clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [clean(v) for v in o]
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating, float)):
        return None if (math.isnan(float(o)) or math.isinf(float(o))) else float(o)
    if isinstance(o, (np.bool_,)):
        return bool(o)
    if isinstance(o, (pd.Timestamp, datetime)):
        return str(o)
    return o


payload = json.dumps(clean(DATA), ensure_ascii=False, separators=(",", ":"), allow_nan=False)
# last-line PII scan of the serialised payload
_bad = re.findall(r"\b03\d{9}\b|\b\+?92\d{10}\b|[\w.]+@[\w.]+\.\w+", payload)
_bad = [b for b in _bad if "example" not in b]
if _bad:
    sys.exit(f"PII pattern found in payload ({_bad[:3]}...): build stopped.")
say(f"Payload   : {len(payload) / 1024:.0f} KB JSON, {sum(1 for s in SEG_META if s['ok'])} segments, {len(PANELS)} panels")
say(f"QA        : {n_flag_total} record(s) flagged to verify across {len(qa_rows)} checks")
if DUMP:
    Path(DUMP).write_text(payload, encoding="utf-8")
if CHECK_ONLY:
    say("--check: nothing written.")
    LOG.write_text("\n".join(LOGLINES), encoding="utf-8")
    sys.exit(0)

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITER = 250_000
salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(PASSWORD.encode("utf-8"))
ct = AESGCM(key).encrypt(iv, zlib.compress(payload.encode("utf-8"), 9), None)
b64 = lambda b: base64.b64encode(b).decode("ascii")
ENC = json.dumps({"v": 1, "iter": ITER, "salt": b64(salt), "iv": b64(iv), "ct": b64(ct)})

tpl = TPL.read_text(encoding="utf-8")
assert "__ENCRYPTED__" in tpl, "template placeholder missing"
html = (tpl.replace("__ENCRYPTED__", ENC)
           .replace("__PROJECT_TITLE__", CFG["title"])
           .replace("__DOMAIN__", CFG.get("domain", "")))
OUT.write_text(html, encoding="utf-8")
say(f"Wrote     : {OUT.name} ({len(html) / 1024:.0f} KB, payload encrypted with AES-256-GCM)")
say(f"Built     : {META['built']}")
LOG.write_text("\n".join(LOGLINES), encoding="utf-8")
