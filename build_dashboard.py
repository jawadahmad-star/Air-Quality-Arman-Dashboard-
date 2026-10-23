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
import difflib
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
sys.path.insert(0, str(HERE))
import dashboard_text as TX
ROOT = HERE.parent
DATA_DIR = Path(os.environ["AQP_DATA_DIR"]) if os.environ.get("AQP_DATA_DIR") else HERE          # data = the .dta files sitting in this Dashboard folder
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
PASSWORD = os.environ.get("AQP_PASSWORD") or CFG["password"]      # env override is for local test builds only
MIN_DUR = float(CFG.get("min_duration_min", 15))
MAX_DUR = float(CFG.get("max_duration_min", 120))
FIELD_HOURS = tuple(CFG.get("field_hours", [7, 21]))
MIN_CELL = int(CFG.get("small_cell_min", 5))
PREFER = CFG.get("prefer", "dta")
ANON_ENUM = bool(CFG.get("anonymise_enumerators", True))
CLASS_THRESHOLD = float(CFG.get("class_threshold", 0.40))      # overridden below by the rule actually coded in the form
RAW_INCLUDE_PII = bool(CFG.get("raw_download_include_pii", True))    # raw-data download: every column as delivered; false strips identifier columns
RAW_DOWNLOAD = bool(CFG.get("raw_download", True))
# don't-know / non-response flags (client comment, Sep 2026): a question is flagged when more than
# this share of the respondents it was actually asked answer Don't Know / Refused / Not sure or leave
# it blank; a completed record is flagged when this many of its don't-know-eligible answers land there.
DK_Q_THRESHOLD = float(CFG.get("dk_question_threshold_pct", 5.0))
DK_RECORD_MIN = int(CFG.get("dk_record_min_count", 4))
# standard South Asian field convention: a respondent (or enumerator, to satisfy a "required" constraint)
# types a round, implausible number into an otherwise-numeric field instead of a formal Don't-Know option.
# 999 is the classic one (client comment, Sep 2026: 35/36 "grade_exam" answers were literal "999"); the
# others are the common siblings checked for in the same pass. Extend this list, don't hardcode a new one.
DK_SENTINEL_VALUES = set(CFG.get("dk_sentinel_values", [999, 9999, -999, 998, 8888]))

# columns that identify a person, a household or a device: dropped on load
PII_EXACT = {"fname", "lname", "address", "phone_mobile", "phone_landline", "mobile_number", "name_child",
             "school_child", "deviceid", "subscriberid", "simid", "devicephonenum", "username", "caseid",
             "text_audit_file", "enum_other", "tuition_type", "air_school_which", "long_term_medical_o",
             "dec_maker_o", "gender_o", "primary_source_o", "measure_air_pollution_o", "actions_protect_c_o",
             "travel_school_o", "more_pay_o", "text_audit", "receiver_name", "child_resp_ill_o", "long_term_medical_o"}
PII_PREFIX = ("geo_", "gps", "phone", "mobile_number")


# ====================================================================== load
ORIG_NAMES = {}


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
    _orig = list(df.columns)
    df.columns = [_norm_col(c) for c in df.columns]
    for _n, _o in zip(df.columns, _orig):
        ORIG_NAMES.setdefault(_n, str(_o))           # remember the export's own column names for the raw-data download
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    out = pd.DataFrame({c: _stringify(df[c]) for c in df.columns})
    for c in ("starttime", "endtime", "SubmissionDate"):
        if c in out:
            out[c] = pd.to_datetime(out[c], errors="coerce", format="mixed").dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
    out["_key"] = _make_key(out)
    return out


def _make_key(df):
    # SurveyCTO's CSV export names this column "KEY"; Stata lowercases variable
    # names on import/export, so the .dta carries it as "key" - match case-
    # insensitively or the CSV and DTA rows for the same submission never join.
    kcol = next((c for c in df.columns if c.upper() == "KEY"), None)
    if kcol and (df[kcol] != "").any():
        return df[kcol]
    g = lambda c: df[c] if c in df else pd.Series("", index=df.index)
    return g("hh_id") + "|" + g("starttime") + "|" + g("enum_name")


def find_sources():
    files = []
    if DATA_DIR.exists():
        for p in DATA_DIR.glob("*.dta"):
            if "call survey" in p.name.lower():
                continue                       # the phone-call .dta is read separately
            if any(part.lower() in ("media", "attachments") for part in p.relative_to(DATA_DIR).parts[:-1]):
                continue                       # SurveyCTO media downloads (sensor streams, text audits), not the survey export
            nm = p.name.lower()
            if nm.startswith(("ta_", "cv_")) or "prefill" in nm or "check" in nm or "master" in nm:
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
RAW_DF = D.copy()                      # every submission exactly as received (before any dashboard-side derivation), for the raw-data download

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

# ---- Don't Know / Refused / Not sure codes, read off each choice list's own labels (never hard-coded -
# the instrument uses 99 on most lists, 98 on some, 7 on "reason", 0 elsewhere - checked systematically
# against the choices sheet rather than assumed). "Not applicable" is a substantive answer, not excluded here.
_DK_LABEL_RE = re.compile(r"don.?t\s*know|refus", re.I)
DK_CODES = {ln: {v for v, l in pairs if _DK_LABEL_RE.search(l)} for ln, pairs in CH.items()}
DK_CODES = {ln: codes for ln, codes in DK_CODES.items() if codes}

# ---- question label lookup (from the survey sheet, for the DK/non-response breakdown table)
_sv = pd.read_excel(form_path, sheet_name="survey").fillna("")
_sv_lab = "label: eng" if "label: eng" in _sv.columns else next(c for c in _sv.columns if str(c).startswith("label"))
QLAB = {}
for _, _r in _sv.iterrows():
    _nm = str(_r["name"]).strip()
    if _nm:
        _lab = re.sub(r"\s+", " ", str(_r[_sv_lab]).replace("\xa0", " ")).strip()
        QLAB[_nm] = (_lab[:87] + "…") if len(_lab) > 88 else _lab


# ---- rules and ranges read straight off the instrument, so the dashboard never drifts from the form
def _form_attr(name, col):
    if col not in _sv.columns:
        return ""
    v = _sv.loc[_sv["name"].astype(str).str.strip() == name, col]
    return str(v.iloc[0]) if len(v) else ""


FORM_VARS = set(_sv["name"].astype(str).str.strip()) - {""}
_thr_src = next((str(r["calculation"]) for _, r in _sv.iterrows() if re.fullmatch(r"\w*percent_parents", str(r["name"]).strip()) and str(r["calculation"]).strip()), "")
_m = re.search(r"\*\s*(0?\.\d+)", _thr_src)
if _m:
    CLASS_THRESHOLD = float(_m.group(1))
THR_PCT = int(round(100 * CLASS_THRESHOLD))
TX.THR_DEFAULT = THR_PCT
PCTS = f"{THR_PCT}%"                                            # e.g. "30%": used in every class-rule label
# schools finished under the earlier 40% protocol keep 40%; every later school uses the form's rule (config "school_thresholds": {school_list code: share})
SCHOOL_THR = {int(k): float(v) for k, v in (CFG.get("school_thresholds") or {}).items()}
_m = re.search(r"<=\s*(\d+)", _form_attr("contribute_will", "constraint"))
CAP = int(_m.group(1)) if _m else 2000                          # highest bid / random price
CAPS = f"{CAP:,}"
_m = re.search(r"round\(\s*(\d+)\s*\*\s*1\.05", _form_attr("tp_amt1", "calculation"))
TP_BASE = int(_m.group(1)) if _m else CAP                       # "today" amount in the time-preference questions
_m = re.search(r"\*\s*(\d+)\s*-\s*sum", _form_attr("topup_amount", "calculation"))
TOPUP_BASE = int(_m.group(1)) if _m else CAP                    # top-up budget per child
_reps, _in_rep = [], False
for _, _r in _sv.iterrows():
    _t = str(_r["type"]).strip()
    if _t == "begin repeat":
        _in_rep = True
    elif _t == "end repeat":
        _in_rep = False
    elif _in_rep and _t.split(" ")[0] not in ("begin", "end", "note") and str(_r["name"]).strip():
        _reps.append(str(_r["name"]).strip())
REPEAT_VARS = _reps                                             # fields that repeat once per child
CLASS_GRADE = {}                                                # which_class_2 code -> school class (4..7)
for _v, _l in CH.get("which_class_2", []):
    _g = re.search(r"\d+", _l)
    if _g:
        CLASS_GRADE[_v] = int(_g.group())
STUDY_GRADES = sorted(CLASS_GRADE.values()) or [4, 5, 6]
SCHOOL_LABEL = dict(CH.get("school_list", []))                  # school_list code -> English school name
say(f"Form rules : class rule {PCTS} of parents · bids up to Rs {CAPS} · time-pref base Rs {TP_BASE:,} · top-up Rs {TOPUP_BASE:,} per child · "
    f"classes {STUDY_GRADES[0]}-{STUDY_GRADES[-1]} · {len(SCHOOL_LABEL)} schools in the form's school list")


def lab_of(ln, code):
    return next((l for v, l in CH.get(ln, []) if v == code), str(code))


# ---------------------------------------------------------------- frame
frame = None
_fc = ([DATA_DIR / "_dummy" / "prefill_dummy.xlsx"] if mode == "demo" else []) + FRAME_CANDIDATES
frame_path = next((p for p in _fc if p.exists()), None)
if frame_path:
    _raw = pd.read_excel(frame_path)
    _raw.columns = [str(c).strip().lower() for c in _raw.columns]
    _f = _raw.iloc[:, :3].copy(); _f.columns = ["hh_id", "order_key", "ap_lm_arm"]
    for _t, _names in (("school", ("school", "school_name", "campus")), ("grade", ("grade", "class", "class_grade")),
                       ("section", ("section", "sec")), ("class_size", ("class_size", "class_n"))):
        _c = next((n for n in _names if n in _raw.columns), None)
        if _c:
            _f[_t] = _raw[_c]
    _f = _f.dropna(subset=["hh_id", "order_key", "ap_lm_arm"])
    _f[["hh_id", "order_key", "ap_lm_arm"]] = _f[["hh_id", "order_key", "ap_lm_arm"]].astype(int)
    frame = _f.drop_duplicates("hh_id").set_index("hh_id")
    say(f"Frame     : {len(frame)} households ({frame_path.name})")
else:
    say("Frame     : none found - targets derived from the data")
HAS_FRAME_CLASS = frame is not None and {"school", "grade"} <= set(frame.columns)
if frame is not None and not HAS_FRAME_CLASS:
    say("Frame     : no school / grade columns - classes are read from the school_child and grade_child text (add school + grade to the frame for reliable classes)")

# ---------------------------------------------------------------- sampling plan: the selected-schools list
# One row per school with the number of sections in each class. This is what defines the classes in the study
# (the prefill file is only an ID pool: hh_id -> arm / question order, with no school information).
_plan_files = []                            # deliberately not used: schools, classes and class sizes are learned from the data as it arrives
_norm_name = lambda x: re.sub(r"[^a-z0-9]", "", str(x).lower())
PLAN = {}                                   # (school name as in the form, class) -> number of sections
if _plan_files:
    _pl = pd.read_excel(_plan_files[-1])
    _pl.columns = [str(c).strip() for c in _pl.columns]
    _ncol = next((c for c in _pl.columns if "name of school" in c.lower()), None)
    _by_norm = {_norm_name(l): l for l in SCHOOL_LABEL.values()}
    for _, _r in _pl.iterrows():
        if _ncol is None or pd.isna(_r[_ncol]):
            continue
        _nm = _by_norm.get(_norm_name(_r[_ncol]))
        if _nm is None:
            say(f"  ! School plan: '{_r[_ncol]}' is not in the form's school_list - kept under its own name")
            _nm = str(_r[_ncol]).strip()
        for _g in STUDY_GRADES:
            _c = next((c for c in _pl.columns if re.fullmatch(rf"(?i)grade\s*{_g}\s*total\s*sections?", c)), None)
            if _c is not None and pd.notna(_r[_c]) and int(_r[_c]) > 0:
                PLAN[(_nm, _g)] = int(_r[_c])
    say(f"School plan: {len({k[0] for k in PLAN})} schools, {sum(PLAN.values())} classes in grades {STUDY_GRADES[0]}-{STUDY_GRADES[-1]} ({_plan_files[-1].name})")
HAS_PLAN = bool(PLAN)


def _norm_sec(sch, gr, sec):
    """section as the plan knows it: "" for a one-section class, otherwise A, B ... (typed A / a / 1 / Sec-B all map)"""
    n = PLAN.get((sch, int(gr))) if pd.notna(gr) else None
    if n == 1:
        return ""
    t = str(sec if pd.notna(sec) else "").strip().upper()
    if not t:
        return ""
    m = re.search(r"\d+", t)
    if m and not re.search(r"[A-Z]", re.sub(r"SEC(TION)?", "", t)):
        i = int(m.group())
        return chr(64 + i) if 1 <= i <= 26 else t
    m = re.fullmatch(r"(?:SEC(?:TION)?[\s\-_:]*)?([A-Z])", t)
    return m.group(1) if m else t


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

# ---- school and class of each household: from the frame, else from what the enumerator typed
_short = lambda x: re.sub(r"^(govt\.?|government)\s+", "", str(x).strip(), flags=re.I)
SCHOOL_THR_BY_NAME = {_short(SCHOOL_LABEL[c]): t for c, t in SCHOOL_THR.items() if c in SCHOOL_LABEL and abs(t - CLASS_THRESHOLD) > 1e-9}
_old_pcts = sorted({int(round(100 * t)) for t in SCHOOL_THR_BY_NAME.values()})
_old_names = " & ".join(re.sub(r"^(GGHS|GHS)\s+", "", n).title() for n in SCHOOL_THR_BY_NAME)
PCTM = PCTS if not SCHOOL_THR_BY_NAME else f"{PCTS} ({' / '.join(f'{x}%' for x in _old_pcts)} in {_old_names}, finished under the earlier protocol)"
PCTM_SHORT = PCTS if not SCHOOL_THR_BY_NAME else f"{PCTS} / {' / '.join(f'{x}%' for x in _old_pcts)}"


def _norm_school_names(series):
    """Collapse free-typed school names that are the same school spelled/cased differently
    (e.g. "Dharampura campus" / "Dharumpura campus" / "Governmeny boys school dhrampura")
    into one canonical label, without assuming how many distinct schools there are."""
    s = series.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    s = s.where(s != "", np.nan)
    freq = s.dropna().value_counts()
    if freq.empty:
        return s

    def _core(x):
        x = x.lower()
        x = re.sub(r"\bgovernment\b|\bgovernmeny\b|\bgovt\.?\b", " ", x)
        x = re.sub(r"\bboys?\b|\bgirls?\b|\bschool\b|\bcampus\b", " ", x)
        return re.sub(r"[^a-z]", "", x)

    clusters = []
    for val in freq.index:
        core = _core(val)
        hit = next((c for c in clusters if difflib.SequenceMatcher(None, core, c["core"]).ratio() >= 0.72), None)
        if hit is None:
            hit = {"core": core, "members": []}
            clusters.append(hit)
        hit["members"].append(val)

    label_of = {}
    for c in clusters:
        members = c["members"]
        proper = [m for m in members if re.fullmatch(r"[A-Z][a-z]+(?: [A-Z][a-z]+)*", m)]
        canon = max(proper, key=lambda m: freq[m]) if proper else max(members, key=lambda m: freq[m])
        for m in members:
            label_of[m] = canon
    return s.map(label_of)


def _grade_of(t):
    m = re.search(r"\d+", str(t)); return float(m.group()) if m else np.nan


def _sec_of(t):
    t = re.sub(r"(?<=\d)(st|nd|rd|th)\b", "", str(t), flags=re.I)          # "5th" is a grade, not section T
    m = re.search(r"\d+\s*[-_ ]?\s*([A-Za-z])", t); return m.group(1).upper() if m else ""


def _ckey_of(sch, gr, sec):
    return f"{str(sch).strip()}|{int(gr)}|{str(sec).strip().upper()}"


def _cls_label(ck):
    sch, gr, sec = str(ck).split("|")
    return f"{_short(sch)} · Class {gr}" + (f"-{sec}" if sec else "")


_txt_sch = _norm_school_names(D["school_child"] if "school_child" in D else pd.Series("", index=D.index))
_txt_gr = D["grade_child"] if "grade_child" in D else pd.Series("", index=D.index)
D["_sch"] = _txt_sch.where(_txt_sch != "", np.nan)
D["_gradeN"] = _txt_gr.map(_grade_of)
D["_sec"] = _txt_gr.map(_sec_of)
if HAS_FRAME_CLASS:
    frame["_gradeN"] = pd.to_numeric(frame["grade"], errors="coerce")
    frame["_sec"] = frame["section"].fillna("").astype(str).str.strip().str.upper() if "section" in frame else ""
    frame["_sch"] = frame["school"].astype(str).str.strip()
    frame["_ckey"] = [_ckey_of(a_, b_, c_) if pd.notna(b_) else np.nan for a_, b_, c_ in zip(frame["_sch"], frame["_gradeN"], frame["_sec"])]
    D["_sch"] = D["_hh"].map(frame["_sch"]).fillna(D["_sch"])
    D["_gradeN"] = D["_hh"].map(frame["_gradeN"]).fillna(D["_gradeN"])
    D["_sec"] = D["_hh"].map(frame["_sec"]).fillna(D["_sec"])
# the enumerator's own record of the sampled class (enum_school / enum_class / enum_section in the form) is the
# authoritative source and beats the frame and the typed child fields wherever it is filled in
_es = num(D, "enum_school").map(lambda v: SCHOOL_LABEL.get(int(v), np.nan) if pd.notna(v) else np.nan)
_ec = num(D, "enum_class").map(lambda v: CLASS_GRADE.get(int(v), np.nan) if pd.notna(v) else np.nan)
_enum_ok = _es.notna() & _ec.notna()
if "enum_school" in D:
    D["_sec"] = D["_sec"].fillna("").astype(object)
    D.loc[_enum_ok, "_sch"] = _es[_enum_ok]
    D.loc[_enum_ok, "_gradeN"] = _ec[_enum_ok].astype(float)
    _esec = D["enum_section"] if "enum_section" in D else pd.Series("", index=D.index)
    D.loc[_enum_ok, "_sec"] = _esec[_enum_ok].fillna("")
    say(f"Class info: {int(_enum_ok.sum())} of {len(D)} submission(s) carry the enumerator's school / class record (enum_school, enum_class)")
D["_sec"] = D["_sec"].fillna("")
D["_sec"] = [_norm_sec(a_, b_, c_) for a_, b_, c_ in zip(D["_sch"], D["_gradeN"], D["_sec"])]
D["_ckey"] = [_ckey_of(a_, b_, c_) if (pd.notna(a_) and pd.notna(b_)) else np.nan for a_, b_, c_ in zip(D["_sch"], D["_gradeN"], D["_sec"])]
D["_cls"] = D["_ckey"].map(lambda k: _cls_label(k) if pd.notna(k) else np.nan)
D["_row"] = np.arange(len(D))

# ---- typed class lists (<project>/New folder/<school>/Typed_Lists/CLASS 4A.xlsx ...): where a school has them, the list decides
# which section each child is in (the enumerator's class / section and the typed grade are often wrong). PII stays local: only the
# resulting class counts reach index.html.
_lnm = lambda x: re.sub(r"\s+", " ", re.sub(r"[^a-z ]", " ", str(x).lower())).strip()
_ldig = lambda x: re.sub(r"\D", "", str(x))[-10:]
LIST_ROWS, LIST_SIZE, LIST_CLASSES = {}, {}, {}           # school label -> [(name, father, grade, sec)] / (school label, grade, sec) -> rows / school label -> {(grade, sec)}
try:
    for _td in sorted(HERE.parent.glob("*/*/Typed_Lists")):
        _tw = {w for w in re.sub(r"[^a-z0-9 ]", " ", _td.parent.name.lower().replace("_", " ")).split() if not w.isdigit() and w not in ("emis", "gghs", "ghs", "school", "girls", "high", "govt", "cdg", "newly", "upgraded")}
        _lab = next((l for l in SCHOOL_LABEL.values() if len(_tw & set(re.sub(r"[^a-z0-9 ]", " ", l.lower()).split())) >= 2), None)
        if not _lab:
            continue
        for _xf in sorted(_td.glob("CLASS *.xlsx")):
            _m = re.search(r"CLASS\s*(\d+)\s*([A-Z]?)", _xf.stem, flags=re.I)
            _raw_l = pd.read_excel(_xf, header=None)
            _hr = next((i for i in range(min(8, len(_raw_l))) if any("student name" in str(v).lower() for v in _raw_l.iloc[i])), None)
            if not _m or _hr is None:
                continue
            _hd = [str(v).strip().lower() for v in _raw_l.iloc[_hr]]
            _cn, _cf = _hd.index("student name"), (_hd.index("father name") if "father name" in _hd else None)
            _cp = [i_ for i_, h_ in enumerate(_hd) if h_.startswith("phone")]
            for _, _lr in _raw_l.iloc[_hr + 1:].iterrows():
                if pd.isna(_lr.iloc[_cn]) or not _lnm(_lr.iloc[_cn]):
                    continue
                LIST_ROWS.setdefault(_lab, []).append((_lnm(_lr.iloc[_cn]), _lnm(_lr.iloc[_cf]) if _cf is not None else "", int(_m.group(1)), _m.group(2).upper(), {_ldig(_lr.iloc[i_]) for i_ in _cp if len(_ldig(_lr.iloc[i_])) == 10}))
        for _r_ in LIST_ROWS.get(_lab, []):
            LIST_SIZE[(_lab, _r_[2], _r_[3])] = LIST_SIZE.get((_lab, _r_[2], _r_[3]), 0) + 1
            LIST_CLASSES.setdefault(_lab, set()).add((_r_[2], _r_[3]))
        say(f"Class lists: {_lab}: {len(LIST_ROWS.get(_lab, []))} children in {len(LIST_CLASSES.get(_lab, ()))} sections ({_td.parent.name})")
except Exception as e:
    say(f"  ! typed class lists skipped: {e!r}")


_LSTOP = {"m", "md", "muhammad", "mohammad", "muhammed", "mohd", "bibi", "begum", "nan", "khan"}
_lrat = lambda a, b: difflib.SequenceMatcher(None, a, b).ratio()
_ltok = lambda x: [t for t in x.split() if t not in _LSTOP]


def _lscore(n, parent, row):
    """how well a typed-list row fits the child: name (also with the father's name written after it) + family evidence"""
    nm, fa = row[0], row[1]
    tn, tl, tf, tp = _ltok(n), _ltok(nm), _ltok(fa), _ltok(parent)
    sim = max(_lrat(n, nm), _lrat(n, nm + " " + fa),
              0.9 if (min(len(n.split()), len(nm.split())) >= 2 and (set(n.split()) <= set(nm.split()) or set(nm.split()) <= set(n.split()))) else 0)
    fam = any(any(_lrat(t, b_) >= 0.8 for b_ in tl[1:] + tf) for t in tn[1:] + tp)
    if sim < 0.85 and tn and tl and _lrat(tn[0], tl[0]) >= 0.8 and fam:
        sim = 0.85
    return sim + (0.1 if fam and sim >= 0.85 else 0)


def _list_match(sch, name, parent, gtxt, phones=(), hint=None, sib=""):
    """(grade, section) of the child in the typed list of its school, or None when not found / ambiguous.
    Parent phone + child name first, then name + family, then the typed grade / enumerator's class to break ties."""
    rows, n, parent = LIST_ROWS.get(sch), _lnm(name), _lnm(parent)
    if not rows or not n:
        return None
    pg = [r for r in rows if r[4] & set(phones)] if phones else []
    if pg:
        ps = sorted(((max(_lrat(n, r[0]), _lrat(n.split()[0], r[0].split()[0]) if r[0] else 0), r) for r in pg), key=lambda x: -x[0])
        if ps[0][0] >= 0.7 and (len(ps) == 1 or ps[0][0] > ps[1][0] + 0.05 or ps[0][1][2:4] == ps[1][1][2:4]):
            return ps[0][1][2:4]
    sc = [(_lscore(n, parent, r), r) for r in rows]
    best = max(s_ for s_, _ in sc)
    if best >= 0.85:
        cand = [r for s_, r in sc if s_ >= best - 0.01]
        if len({r[2:4] for r in cand}) > 1:
            g_ = _grade_of(gtxt)
            cand = [r for r in cand if r[2] == g_] or cand
        if len({r[2:4] for r in cand}) > 1 and hint:
            cand = [r for r in cand if r[2:4] == hint] or cand
        if len({r[2:4] for r in cand}) == 1:
            return cand[0][2:4]
    # not found by name: an explicit "6th B" typed by the enumerator, if that section exists in the school's list
    t_ = re.sub(r"(?<=\d)(st|nd|rd|th)\b", "", str(gtxt), flags=re.I)
    m_ = re.fullmatch(r"\s*(\d+)\s*[-_ ]?\s*([A-Za-z])\s*", t_)
    if m_ and (int(m_.group(1)), m_.group(2).upper()) in LIST_CLASSES.get(sch, ()):
        return (int(m_.group(1)), m_.group(2).upper())
    # looser, within the class the enumerator named (typed grade, else the class selected at the start of the survey), looking at
    # all its sections: first name roughly the same (spelling is hand-copied) + the surname / father's name / mother's surname fits
    tn = _ltok(n)
    for g_ in dict.fromkeys([x for x in (_grade_of(gtxt), hint[0] if hint else np.nan) if x in STUDY_GRADES]):
        fit = []
        for r in rows:
            if r[2] != g_ or not tn:
                continue
            tl, tf, tp = _ltok(r[0]), _ltok(r[1]), _ltok(parent + " " + sib)      # sib: the other children's names (a surname the parent did not give)
            f_ = _lrat(tn[0], tl[0]) if tl else 0
            fam = any(any(_lrat(t, b_) >= 0.75 for b_ in tl[1:] + tf) for t in tn[1:] + tp)
            if f_ >= 0.7 and fam:
                fit.append((f_, r))
        if fit:
            top = max(f for f, _ in fit)
            c_ = {r[2:4] for f, r in fit if f >= top - 0.05}
            if len(c_) == 1:
                return next(iter(c_))
    # still nothing: the class and section the enumerator selected for this sampled child
    if hint and hint in LIST_CLASSES.get(sch, ()):
        return hint
    return None


LIST_MATCH = {}                                           # (row, child no.) -> (grade, section)
LIST_MISS = []                                            # children of a list school who are not in its list
if LIST_ROWS:
    _lpar = (D["fname"].fillna("").astype(str) + " " + D["lname"].fillna("").astype(str)) if "fname" in D else pd.Series("", index=D.index)
    for _i in range(len(D)):
        _lab = D["_sch"].iat[_i]
        if _lab not in LIST_ROWS:
            continue
        for _k in range(0, 21):
            _sf = f"_{_k}" if _k else ""
            if f"name_child{_sf}" not in D.columns:
                continue
            _nmv = D[f"name_child{_sf}"].iat[_i]
            if not _lnm(_nmv):
                continue
            _gt = D[f"grade_child{_sf}"].iat[_i] if f"grade_child{_sf}" in D.columns else ""
            _ph = {_ldig(D[c_].iat[_i]) for c_ in ("phone_mobile", "mobile_number") if c_ in D.columns}
            _hint = (int(D["_gradeN"].iat[_i]), str(D["_sec"].iat[_i]).strip().upper()) if pd.notna(D["_gradeN"].iat[_i]) else None
            _mt = _list_match(_lab, _nmv, _lpar.iat[_i], _gt, _ph, _hint if _k <= 1 else None,
                             " ".join(_lnm(D[f"name_child_{j_}"].iat[_i]) for j_ in range(1, 21) if j_ != _k and f"name_child_{j_}" in D.columns))   # the enumerator's class is the sampled (first) child's
            if _mt:
                LIST_MATCH[(_i, _k or 1)] = _mt
            else:
                LIST_MISS.append((int(D["_hh"].iat[_i]) if pd.notna(D["_hh"].iat[_i]) else None, str(_nmv)))
    # household-level class: the child in the sampled class if there is one, else the first listed child
    for _i in range(len(D)):
        _ms = [LIST_MATCH[(_i, _k)] for _k in range(1, 21) if (_i, _k) in LIST_MATCH]
        if not _ms or D["_sch"].iat[_i] not in LIST_ROWS:
            continue
        _hit = next((m_ for m_ in _ms if m_[0] == D["_gradeN"].iat[_i]), _ms[0])
        D.iat[_i, D.columns.get_loc("_gradeN")] = float(_hit[0])
        D.iat[_i, D.columns.get_loc("_sec")] = _hit[1]
    D["_ckey"] = [_ckey_of(a_, b_, c_) if (pd.notna(a_) and pd.notna(b_)) else np.nan for a_, b_, c_ in zip(D["_sch"], D["_gradeN"], D["_sec"])]
    D["_cls"] = D["_ckey"].map(lambda k: _cls_label(k) if pd.notna(k) else np.nan)
    say(f"Class lists: {len(LIST_MATCH)} child record(s) placed in their section from the typed lists, {len(LIST_MISS)} not found in a list (not counted in any class), hh_id: {[h_ for h_, _ in LIST_MISS]}")

# ---- repeat group: a wide export carries the per-child fields as <name>_1, <name>_2 ...
K_MAX = 0
for _k in range(1, 21):
    if any(f"{_v}_{_k}" in D.columns and f"{_v}_{_k}" not in FORM_VARS for _v in ("name_child", "grade_child", "final_bid", "contribute_will")):
        K_MAX = _k
REPEAT_LAYOUT = K_MAX >= 1
say(f"Children   : " + (f"repeat-group layout detected, up to {K_MAX} child record(s) per household" if REPEAT_LAYOUT else "one child record per household (no repeat suffixes in this export)"))
D = D.copy()

# ---- PII guard: drop, then assert
def is_pii(c):
    """identifier column, including repeat-group copies such as name_child_1 / school_child_2"""
    base = re.sub(r"(_\d+)+$", "", str(c))
    return (c in PII_EXACT or base in PII_EXACT or str(c).lower().startswith(PII_PREFIX)) and not str(c).startswith("_")


# one-way respondent key (name hash) so duplicates can be told from different parents sharing a pool hh_id; the names themselves are dropped below
import hashlib
_rn = lambda c: D[c].fillna("").astype(str).str.lower().str.strip() if c in D else ""
D["_resp"] = (_rn("fname") + "|" + _rn("lname")).map(lambda x: hashlib.sha1(x.encode("utf-8")).hexdigest()[:10])
_drop = [c for c in D.columns if is_pii(c)]
D = D.drop(columns=_drop)
assert not [c for c in D.columns if is_pii(c)], "PII column survived the guard"

if len(D) and D["_dt"].isna().all():
    say("WARNING: starttime could not be parsed - dates will be empty")


# ---- one record per child: a parent with two children in the school bids once for each, and each bid counts in
# that child's own class. Built from the suffixed repeat columns (or the single unsuffixed set in a no-repeat export).
def _bid_and_draw(df, sfx):
    g = lambda v: num(df, v + sfx)
    chg = (g("change_pay") == 1) & g("change_pay_w").notna()
    bid = pd.Series(np.where(chg, g("change_pay_w"), g("contribute_will")), index=df.index)
    fb, fr = g("final_bid"), g("final_rand_number")
    bid = bid.where(fb.isna(), fb)
    r1_ = g("change_rand_number")
    rand = pd.Series(np.where(fr.notna(), fr, np.where(chg & r1_.notna(), r1_, g("main_contribute_rand_number"))), index=df.index)
    return bid, rand


_kid_parts = []
for _k in range(1, max(K_MAX, 1) + 1):
    _sfx = f"_{_k}" if REPEAT_LAYOUT else ""
    _bid, _rand = _bid_and_draw(D, _sfx)
    _tp = num(D, "total_parents" + _sfx).fillna(num(D, "total_parents_2" + _sfx))
    _gtxt = D["grade_child" + _sfx] if ("grade_child" + _sfx) in D else pd.Series("", index=D.index)
    _g = _gtxt.map(_grade_of)
    _g = _g.where(_g.isin(STUDY_GRADES) | (not HAS_PLAN), np.nan)
    _same = _g.isna() | (_g == D["_gradeN"])
    _g = _g.fillna(D["_gradeN"]) if _k == 1 else _g                                  # first child: the sampled class if not typed
    _sec = np.where(_same, D["_sec"], _gtxt.map(_sec_of))                           # another class of the school: only a typed section is known
    _present = _bid.notna() | _tp.notna() | (_gtxt.astype(str) != "")
    _p = pd.DataFrame({"_row": D["_row"], "_hh": D["_hh"], "_k": _k, "_sch": D["_sch"], "_gradeN": _g, "_sec": _sec,
                       "_bid": _bid, "_rand": _rand, "_tpar": _tp})[_present]
    _kid_parts.append(_p)
KID = pd.concat(_kid_parts, ignore_index=True)
KID["_sec"] = [_norm_sec(a_, b_, c_) for a_, b_, c_ in zip(KID["_sch"], KID["_gradeN"], KID["_sec"])]
# a child typed into another class of the school usually has no section recorded: if the data shows exactly one
# section for that school + grade, the blank belongs to it (otherwise it would show up as a phantom extra class)
_known_secs = {}
for _s_, _g_, _x_ in zip(pd.concat([D["_sch"], KID["_sch"]]), pd.concat([D["_gradeN"], KID["_gradeN"]]), pd.concat([D["_sec"], KID["_sec"]])):
    if pd.notna(_s_) and pd.notna(_g_) and str(_x_).strip():
        _known_secs.setdefault((_s_, _g_), set()).add(str(_x_).strip())
KID["_sec"] = [(next(iter(_known_secs[(a_, b_)])) if (not str(c_).strip() and (a_, b_) in _known_secs and len(_known_secs[(a_, b_)]) == 1) else c_)
               for a_, b_, c_ in zip(KID["_sch"], KID["_gradeN"], KID["_sec"])]
KID["_ckey"] = [_ckey_of(a_, b_, c_) if (pd.notna(a_) and pd.notna(b_)) else np.nan for a_, b_, c_ in zip(KID["_sch"], KID["_gradeN"], KID["_sec"])]
if LIST_ROWS:                                             # typed lists decide the class of every child in their schools
    _kg, _ks = [], []
    for _r_, _k_, _s_, _g_, _x_ in zip(KID["_row"], KID["_k"], KID["_sch"], KID["_gradeN"], KID["_sec"]):
        if _s_ in LIST_ROWS:
            _mt = LIST_MATCH.get((int(_r_), int(_k_)))
            _kg.append(float(_mt[0]) if _mt else np.nan); _ks.append(_mt[1] if _mt else "")
        else:
            _kg.append(_g_); _ks.append(_x_)
    KID["_gradeN"], KID["_sec"] = _kg, _ks
    KID["_ckey"] = [_ckey_of(a_, b_, c_) if (pd.notna(a_) and pd.notna(b_)) else np.nan for a_, b_, c_ in zip(KID["_sch"], KID["_gradeN"], KID["_sec"])]
_ok_k = KID["_rand"].notna() & KID["_bid"].notna()
KID["_hit"] = np.where(_ok_k, (KID["_rand"] <= KID["_bid"]).astype(float), np.nan)
KID["_pclear"] = np.where(KID["_bid"].notna(), (KID["_bid"].clip(0, CAP) + 1) / (CAP + 1.0), np.nan)

# first child's answers become the household-level columns the rest of the dashboard reads
SMULTI = {str(n).strip() for n, t in zip(_sv["name"], _sv["type"]) if str(t).startswith("select_multiple")}
if REPEAT_LAYOUT:
    _aliased = 0
    for _v in REPEAT_VARS:
        _src = f"{_v}_1"
        if _src in D.columns and (_v not in SMULTI or any(c.startswith(_src + "_") for c in D.columns)):
            D[_v] = D[_src]
            _aliased += 1
    D = D.copy()
    say(f"Children   : {_aliased} repeat-group field(s) read from child 1 for the household-level tabs; {len(KID)} child record(s) for the class counts")

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
# hh_id comes from a shared ID pool (5,000 IDs), so the same hh_id appears for different parents across rounds:
# a duplicate is the same hh_id AND the same respondent name, not the hh_id alone.
n_dup_rows = 0                                           # every completed submission counts: the same pool hh_id is reused by different parents
C = comp_all.sort_values("_dt").reset_index(drop=True)

KIDC = KID[KID["_row"].isin(C["_row"])].copy()           # child records belonging to the analysis-base interviews

# ---- the classes of the study and how many parents each one has
# With the selected-schools list the study is a fixed set of classes; the prefill file is only an ID pool (its
# 5,000 IDs are not the sample). Class size is what the interviews report as "total parents in the class";
# until a class has reported, the planning size from the config is used.
UNIVERSE = {}                                             # class key -> in the plan?
if HAS_PLAN:
    for (_s, _g), _n in PLAN.items():
        for _sc in ([""] if _n == 1 else [chr(65 + _i) for _i in range(_n)]):
            UNIVERSE[_ckey_of(_s, _g, _sc)] = True
if not HAS_PLAN:                                          # no plan: classes are whatever the data shows
    for _ck in set(D["_ckey"].dropna()) | set(KIDC["_ckey"].dropna()):
        UNIVERSE.setdefault(_ck, False)
N_UNMATCHED = int((~KIDC["_ckey"].isin(UNIVERSE.keys())).sum()) if HAS_PLAN else 0     # child records whose school / class / section is not in the plan
if N_UNMATCHED:
    say(f"  ! Class match: {N_UNMATCHED} child record(s) could not be matched to a planned class (grade or section missing / not in the plan) - not counted in any class")


_nkc = lambda x: re.sub(r"\W+", " ", str(x)).upper().strip()
# ---- per-school "Parents_Coverage_*.xlsx" sheets (list-checked parent counts): when one sits in this folder its numbers replace the
# survey-derived class sizes / outcomes for that school: actual parents, complete, and the call outcomes of everyone else
COV_CLASS, COV_SCHOOL = {}, {}
cls_schools_all = sorted({_short(_k.split("|")[0]) for _k in UNIVERSE} | {_short(l) for l in SCHOOL_LABEL.values()})
try:
    import openpyxl
    for _cf in sorted(HERE.glob("Parents_Coverage_*.xlsx")):
        _ws = openpyxl.load_workbook(_cf, data_only=True)["Coverage"]
        _title = str(_ws["A1"].value or "").split("(EMIS")[0]
        _tw = set(_nkc(_title).split()) - {"GGHS", "GHS", "SCHOOL", "GIRLS", "HIGH", "GOVT", "CDG"}
        _sn = next((_s for _s in cls_schools_all if len(_tw & set(_nkc(_s).split())) >= 2), None)
        if not _sn:
            say(f"  ! {_cf.name}: no matching school in the form, skipped")
            continue
        _tot = {"actual": 0, "complete": 0, "refused": 0, "off": 0, "notpicked": 0, "wrong": 0, "invalid": 0, "sibling": 0, "final": 0}
        for _r in _ws.iter_rows(min_row=4, values_only=True):
            if not isinstance(_r[0], (int, float)):
                break
            _c = {"actual": _r[5], "complete": _r[6], "refused": _r[7], "off": _r[8], "notpicked": _r[9], "wrong": _r[10], "invalid": _r[11], "sibling": _r[12]}
            _c = {k: int(v or 0) for k, v in _c.items()}
            _c["final"] = _c["refused"] + _c["wrong"] + _c["invalid"]
            COV_CLASS[(_sn, int(_r[0]))] = _c
            for k in _tot:
                _tot[k] += _c[k]
        COV_SCHOOL[_sn] = _tot
        say(f"Coverage  : {_sn}: {_tot['actual']} actual parents, {_tot['complete']} complete ({_cf.name})")
except Exception as e:
    say(f"  ! coverage sheets skipped: {e!r}")


def _class_size(ck):
    _ls = LIST_SIZE.get((str(ck).split("|")[0], int(str(ck).split("|")[1]), str(ck).split("|")[2]))
    if _ls:                                              # typed class list: its rows are the parents of the section
        return _ls
    _cv = COV_CLASS.get((_short(str(ck).split("|")[0]), int(str(ck).split("|")[1])))
    if _cv:                                              # list-checked sheet: actual parents in the class
        return _cv["actual"]
    _s = KIDC.loc[KIDC["_ckey"] == ck, "_tpar"].dropna()
    if len(_s):
        return int(_s.median())
    if HAS_FRAME_CLASS:
        _fs = frame.loc[frame["_ckey"] == ck, "class_size"].dropna() if "class_size" in frame else []
        if len(_fs):
            return int(np.median(_fs))
        return int((frame["_ckey"] == ck).sum())
    return int((KIDC["_ckey"] == ck).sum())              # nothing reported yet: no more than the children seen so far


if UNIVERSE and not HAS_FRAME_CLASS:                      # no fixed sample: the target is the parents of the classes seen so far and grows as new classes appear
    N_TARGET_PLAN = sum(_class_size(_ck) for _ck in UNIVERSE)
elif HAS_FRAME_CLASS:
    N_TARGET_PLAN = len(frame)
else:
    N_TARGET_PLAN = len(frame) if frame is not None else max(len(C), 1)
N_TARGET = int(CFG.get("target_override") or N_TARGET_PLAN)
FRAME_IS_POOL = frame is not None and ("enum_school" in D) and not HAS_FRAME_CLASS     # prefill = ID pool, not the list of households to visit
N_ATT = len(D)
N_C = len(C)
say(f"Attempts  : {N_ATT}   completed (analysis base): {N_C}   target: {N_TARGET}")
say(f"Excluded  : {n_out_of_frame} completed row(s) outside the frame, {n_dup_rows} duplicate household row(s)")

# ---- derived analysis columns on C
C["_bid0"] = num(C, "contribute_will")
_chg = (num(C, "change_pay") == 1) & num(C, "change_pay_w").notna()
C["_bid"] = np.where(_chg, num(C, "change_pay_w"), C["_bid0"])
_fb, _fr = num(C, "final_bid"), num(C, "final_rand_number")          # the questionnaire's own final bid / final draw (used when present)
C["_bid"] = np.where(_fb.notna(), _fb, C["_bid"])
C["_bid"] = pd.to_numeric(C["_bid"], errors="coerce")
C["_bonus"] = num(C, "contribute_will_bonus")
# the draw that decides whether a parent contributes: final_bid vs final_rand_number from the form;
# for older exports without them, the revised bid meets the new random number, otherwise the original bid meets the first one
_r0, _r1 = num(C, "main_contribute_rand_number"), num(C, "change_rand_number")
_use_new = _chg & _r1.notna()
C["_rand"] = np.where(_fr.notna(), _fr, np.where(_use_new, _r1, _r0))
C["_bid_d"] = C["_bid"]
C["_bid_d"] = pd.to_numeric(C["_bid_d"], errors="coerce")
_ok = C["_rand"].notna() & C["_bid_d"].notna()
C["_paid"] = np.where(_ok, np.where(C["_rand"] <= C["_bid_d"], C["_rand"], 0), np.nan)
C["_hit"] = np.where(_ok, (C["_rand"] <= C["_bid_d"]).astype(float), np.nan)
C["_pclear"] = np.where(C["_bid_d"].notna(), (C["_bid_d"].clip(0, CAP) + 1) / (CAP + 1.0), np.nan)   # chance a random price clears this bid
C["_income"] = num(C, "hh_income")
C["_age"] = num(C, "age")
C["_hhsize"] = num(C, "hh_member")

C["_grade"] = C["_gradeN"]
_ex = C["grade_exam"].astype(str).str.extract(r"(\d+(?:\.\d+)?)")[0] if "grade_exam" in C else pd.Series(np.nan, index=C.index)
C["_exam"] = pd.to_numeric(_ex, errors="coerce").where(lambda s: (s >= 0) & (s <= 100))

# implied annual discount rate from the bisection end-node
def _tp_amt(node):
    node = float(np.clip(node, 1, 31))
    return round(TP_BASE * 1.05 * math.pow(3.809524, (node - 1) / 30) / 50) * 50


C["_tpamt"] = num(C, "tp_bin").map(lambda v: np.nan if pd.isna(v) else _tp_amt(v))
C["_disc"] = (C["_tpamt"] / float(TP_BASE) - 1.0) * 100.0
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

for _gv in sorted(C["_gradeN"].dropna().unique()):
    SEG_DEFS.append((f"cls{int(_gv)}", f"Class {int(_gv)}", "Class grade", (lambda vv: lambda d: d["_gradeN"] == vv)(_gv)))
_schools = sorted(C["_sch"].dropna().unique())
for _i, _sn in enumerate(_schools, start=1):
    SEG_DEFS.append((f"sch{_i}", _short(_sn), "School", (lambda nm: lambda d: d["_sch"] == nm)(_sn)))

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
        return {"items": items(labels, cnt, base), "base": base, "hist": True}
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


# ====================================================================== don't-know / non-response checks
# Client comment (Sep 2026): flag a question when Don't Know / Refused / non-response runs high among
# the people it was actually asked of, and flag a whole interview when several of its answers land there.
# Two sources feed this, both read from the instrument rather than hard-coded:
#   1. DK_QUESTIONS - select_one/select_multiple questions where the instrument itself gives a DK/Refused/
#      Not-sure option (via DK_CODES, read off the choices sheet - see above).
#   2. SENTINEL_QUESTIONS (further down) - required text/integer/decimal questions with NO such option,
#      where a respondent can only signal "don't know" by typing an implausible/sentinel number (e.g. 999)
#      into an otherwise-numeric field. Added Sep 2026 after grade_exam turned up 97% literal "999".
# Two of the DK_QUESTIONS below are only relevant to a subset of respondents (sickness_air, more_pay) -
# relevance is evaluated the same way the form does, so a skip-pattern "not applicable" case never counts.
def _rel_all(df):
    return pd.Series(True, index=df.index)


def _rel_sickness_air(df):
    codes = [v for v, _ in CH.get("long_term_med", [])]
    toks, _ = _tokens(df, "long_term_medical", codes)
    return ~toks.map(lambda t: 7 in t)                       # asked unless "None of the above" was ticked


def _rel_more_pay(df):
    return (num(df, "contribute_will") != CAP) & (num(df, "change_pay") == 2)


DK_QUESTIONS = [
    # (column, choice list, "one"/"multi", relevance fn matching the form's own skip logic)
    ("sickness_air", "yesno_dk", "one", _rel_sickness_air),
    ("beleifs_air_pollute", "belief_air", "multi", _rel_all),
    ("check_air_quality", "check_air", "one", _rel_all),
    ("pollute_believe", "belief", "one", _rel_all),
    ("risk_air_child", "disease_risk", "one", _rel_all),
    ("measure_air_pollution", "measures", "multi", _rel_all),
    ("actions_protect_c", "measures", "multi", _rel_all),
    ("more_pay", "reason", "one", _rel_more_pay),
    ("certain_choice", "choice", "one", _rel_all),
    ("tp1", "timeopt", "one", _rel_all),
    ("tp2", "timeopt", "one", _rel_all),
    ("tp3", "timeopt", "one", _rel_all),
    ("tp_check", "timeopt", "one", _rel_all),
]
DK_QUESTIONS = [q for q in DK_QUESTIONS if q[1] in DK_CODES]
# tp1/tp2/tp3/tp_check share one templated label (the amount is a calculated field) - tell them apart here
DK_LABEL_OVERRIDE = {
    "tp1": f"Time-preference bisection - step 1 (Rs {TP_BASE:,} today vs. more in a year)",
    "tp2": "Time-preference bisection - step 2",
    "tp3": "Time-preference bisection - step 3",
    "tp_check": f"Time-preference consistency check (Rs {TP_BASE:,} today vs. Rs {TP_BASE:,} in a year)",
}

# Free-text/integer/decimal "required" questions that carry no formal DK/Refused choice option at all, so
# the DK_CODES scan above never sees them - a respondent (or an enumerator satisfying SurveyCTO's
# required=yes) can only signal "don't know" here by typing a sentinel value or leaving a value that's
# implausible for the question. Found by systematically scanning every text/integer/decimal question in the
# survey sheet that is required=yes and has no DK_CODES list (~50 of them: age, hh_income, hh_member,
# time_reach, grade_child, the WTP/classroom-logistics fields, etc.) against the real pilot data for
# round-number heaping (client comment, Sep 2026: "grade_exam" was 35/36 = 97% literal "999"). Everything
# else checked came back clean - re-run this scan against each fresh data drop and add anything new here;
# do not hardcode "999" anywhere else in the script, it lives in DK_SENTINEL_VALUES above.
SENTINEL_QUESTIONS = [
    # (column, plausible (lo, hi) range for a genuine answer, relevance fn)
    ("grade_exam", (0, 100), _rel_all),
]


def _numeric_val(s):
    ex = s.astype(str).str.extract(r"(-?\d+(?:\.\d+)?)")[0]
    return pd.to_numeric(ex, errors="coerce")


_dk_cols = [q[0] for q in DK_QUESTIONS] + [q[0] for q in SENTINEL_QUESTIONS if q[0] in C]
dk_q_rows = []
_dk_flag = pd.DataFrame(False, index=C.index, columns=_dk_cols)
_dk_elig = pd.DataFrame(False, index=C.index, columns=_dk_cols)
for qname, lst, kind, relfn in DK_QUESTIONS:
    rel = relfn(C).fillna(False)
    codes = DK_CODES[lst]
    if kind == "one":
        v = num(C, qname)
        is_dk = v.isin(codes)                                    # blanks are not counted: skip logic leaves them legitimately empty
    else:
        _codes_all = [c for c, _ in CH.get(lst, [])]
        toks, _ = _tokens(C, qname, _codes_all)
        is_dk = toks.map(lambda t: bool(t & codes))
    is_dk = is_dk & rel
    _dk_elig[qname] = rel
    _dk_flag[qname] = is_dk
    base = int(rel.sum())
    n_dk = int(is_dk.sum())
    p_dk = pct(n_dk, base) if base else None
    dk_q_rows.append({"q": qname, "label": DK_LABEL_OVERRIDE.get(qname, QLAB.get(qname, qname)), "list": lst, "base": base, "n": n_dk, "p": p_dk,
                      "flag": bool(base and p_dk is not None and p_dk > DK_Q_THRESHOLD)})
    if base == 0:
        say(f"  ! DK/NR check: '{qname}' had no relevant respondents - check its relevance rule / column name")

for qname, (lo, hi), relfn in SENTINEL_QUESTIONS:
    if qname not in C:
        say(f"  ! DK/NR sentinel check: '{qname}' not found in this data drop - skipped")
        continue
    rel = relfn(C).fillna(False)
    v = _numeric_val(C[qname])
    is_dk = v.isin(DK_SENTINEL_VALUES) | (v < lo) | (v > hi)
    is_dk = is_dk & rel
    _dk_elig[qname] = rel
    _dk_flag[qname] = is_dk
    base = int(rel.sum())
    n_dk = int(is_dk.sum())
    p_dk = pct(n_dk, base) if base else None
    dk_q_rows.append({"q": qname, "label": DK_LABEL_OVERRIDE.get(qname, QLAB.get(qname, qname)),
                      "list": f"sentinel ({lo}-{hi} plausible)", "base": base, "n": n_dk, "p": p_dk,
                      "flag": bool(base and p_dk is not None and p_dk > DK_Q_THRESHOLD)})
    if base == 0:
        say(f"  ! DK/NR sentinel check: '{qname}' had no relevant respondents - check its relevance rule / column name")

n_dk_q_total = len(dk_q_rows)
n_dk_q_flagged = sum(1 for r in dk_q_rows if r["flag"])
dk_q_flagged_labels = [r["label"] for r in dk_q_rows if r["flag"]]

# record-level: how many of the DK-eligible questions asked of this respondent came back Don't Know / Refused / Not sure.
# Threshold (DK_RECORD_MIN, default 4 of up to 13 eligible items = roughly a third): an isolated "don't
# know" on one or two items is normal survey noise; a respondent unable or unwilling to substantively
# answer several of the knowledge/opinion items in a row is worth the field team looking again at
# comprehension, rapport, or possible rushing - not a claim that anything was fabricated.
C["_dk_elig_n"] = _dk_elig.sum(axis=1)
C["_dk_n"] = _dk_flag.sum(axis=1)
C["_dk_share"] = np.where(C["_dk_elig_n"] > 0, 100.0 * C["_dk_n"] / C["_dk_elig_n"], 0.0)
n_dk_records = int((C["_dk_n"] >= DK_RECORD_MIN).sum())
say(f"DK/NR checks : {n_dk_q_flagged} of {n_dk_q_total} questions above {DK_Q_THRESHOLD:.0f}% DK/Refused; "
    f"{n_dk_records} of {N_C} completed record(s) with {DK_RECORD_MIN}+ such answers")

# table card shared by Overview (client-visible) and Field Operations (internal): one row per DK-eligible question
_dk_pill = lambda flagged: {"v": 1 if flagged else 0, "t": "Verify" if flagged else "OK", "tone": "amber" if flagged else "good"}
dk_table = {"cols": [{"k": "label", "l": "Question"}, {"k": "list", "l": "Choice list / detection"}, {"k": "base", "l": "Asked", "num": True},
                     {"k": "n", "l": "Don't Know / Refused / Not sure", "num": True}, {"k": "p", "l": "Share", "num": True}, {"k": "flag", "l": "Status", "verdict": True}],
            "rows": [{"label": r["label"], "list": r["list"], "base": r["base"], "n": r["n"],
                      "p": {"v": r["p"] if r["p"] is not None else 0, "t": "—" if r["p"] is None else f"{r['p']:.1f}%"},
                      "flag": _dk_pill(r["flag"])} for r in dk_q_rows],
            "read": (f"<strong>{n_dk_q_flagged}</strong> of {n_dk_q_total} questions are above the {DK_Q_THRESHOLD:.0f}% line, counted among the respondents each question was actually asked of."
                     + (f" Flagged: {', '.join(dk_q_flagged_labels[:4])}{'…' if len(dk_q_flagged_labels) > 4 else ''}." if dk_q_flagged_labels else " None are above the line yet."))}


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
elif FRAME_IS_POOL:                                       # the prefill is an ID pool: the plan is an equal split of the sample over the three arms
    arm_target = {v: int(round(N_TARGET / 3)) for v in (0, 1, 2)}
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

# ---- phone screening calls (Data phone survey): enumerators call parents first, then visit the ones who agree.
# One row per PARENT (a parent called several times counts once; a "Completed" call wins, otherwise the latest call).
# Parents who could not be interviewed by phone (switched off, not picked, wrong number, refused) are added to the
# attempts and the non-response; parents already interviewed in the field are never counted twice.
PHONE_DIR = HERE
PHONE_FORM = ROOT / "Instrument" / "air_quality_call_survey.xlsx"
ph_calls = ph_parents = n_ph_nr = n_ph_agreed = n_ph_refused = n_ph_notreached = n_ph_retry = n_ph_final = 0
ph_disp_n = {}
ph_school = {}                                                   # school -> phone parents not interviewed {nr, retry, final}
_nk = lambda x: re.sub(r"\W+", " ", str(x)).upper().strip()
try:
    _pf = sorted(PHONE_DIR.glob("*Call Survey*.dta"))
    if _pf:
        _p = load_file(_pf[0])
        _p["_hh"] = pd.to_numeric(_p["hh_id"], errors="coerce")
        _p["_st"] = pd.to_numeric(_p["status_survey"], errors="coerce")
        _p["_dt"] = pd.to_datetime(_p["starttime"], errors="coerce")
        _pch = pd.read_excel(PHONE_FORM, sheet_name="choices").fillna("") if PHONE_FORM.exists() else pd.DataFrame(columns=["list_name", "value", "label: eng"])
        PH_ST = {int(r["value"]): str(r["label: eng"]).replace("\u2060", "").strip() for _, r in _pch[_pch["list_name"] == "status_survey"].iterrows()}
        _p = _p.dropna(subset=["_hh", "_st"]).sort_values("_dt")
        ph_calls = len(_p)
        _done_field = set(num(C, "hh_id").dropna().astype(int)) if "hh_id" in C else set(C["_hh"].dropna().astype(int))
        _PHS = {int(r["value"]): _short(str(r["label: eng"])) for _, r in _pch[_pch["list_name"] == "school_list"].iterrows()}
        _fin, _hsch = {}, {}
        for _h, _g in _p.groupby("_hh"):
            _fin[int(_h)] = 1 if (_g["_st"] == 1).any() else int(_g["_st"].iloc[-1])
            _sv = pd.to_numeric(_g["enum_school"], errors="coerce").iloc[-1]
            _hsch[int(_h)] = _PHS.get(int(_sv), "") if pd.notna(_sv) else ""
        ph_parents = len(_fin)
        n_ph_agreed = sum(1 for h, c in _fin.items() if c == 1 and h not in _done_field)
        for h, c in _fin.items():
            if c in (1, 2) or h in _done_field:                  # agreed / partial: not a non-response, and never double-count a field interview
                continue
            ph_disp_n[PH_ST.get(c, str(c))] = ph_disp_n.get(PH_ST.get(c, str(c)), 0) + 1
            _sd = ph_school.setdefault(_hsch.get(h, ""), {"nr": 0, "retry": 0, "final": 0})
            _sd["nr"] += 1
            if c in (7, 8, 10, 11, 12):                          # switched off / no answer / appointment: call again
                n_ph_retry += 1
                _sd["retry"] += 1
            else:                                                # refused, wrong or invalid number, ineligible: no further call
                n_ph_final += 1
                _sd["final"] += 1
            if c in (3, 4, 5, 6):
                n_ph_refused += 1
            elif c in (7, 8, 10, 11, 12, 13, 14):
                n_ph_notreached += 1
        n_ph_nr = sum(ph_disp_n.values())
        say(f"Phone calls: {ph_calls} calls -> {ph_parents} unique parents · {n_ph_nr} unsuccessful · {n_ph_agreed} agreed (shown in the Overview)")
except Exception as e:                                          # never let the phone block kill the build
    say(f"  ! phone-call block skipped: {e!r}")
    ph_disp_n, n_ph_nr, n_ph_refused, n_ph_notreached, n_ph_agreed, n_ph_retry, n_ph_final = {}, 0, 0, 0, 0, 0, 0
# ---- the phone survey is reported on its own tab; the main tabs (attempts, non-response, remaining) are field visits only
PH = {"calls": ph_calls, "parents": ph_parents, "agreed": n_ph_agreed, "nr": n_ph_nr, "retry": n_ph_retry, "final": n_ph_final,
      "refused": n_ph_refused, "notreached": n_ph_notreached, "disp": dict(ph_disp_n), "school": {k: dict(v) for k, v in ph_school.items()}, "rows": []}
try:
    if ph_calls:
        _pr = _p.copy()
        _pr["_sch"] = pd.to_numeric(_pr["enum_school"], errors="coerce").map(_PHS).fillna("—")
        _pr["_cls"] = _pr["enum_class_label"].replace("", "—") if "enum_class_label" in _pr else "—"
        _pr["_en"] = pd.to_numeric(_pr["enum_name"], errors="coerce").map(lambda c: enum_label(c) if pd.notna(c) else "—") if "enum_name" in _pr else "—"
        _pr["_out"] = _pr["_st"].map(lambda c: PH_ST.get(int(c), str(int(c))) if pd.notna(c) else "—")
        _pr["_att"] = _pr.groupby("_hh").cumcount() + 1
        PH["rows"] = [{"d": r["_dt"].strftime("%b %d") if pd.notna(r["_dt"]) else "—", "e": r["_en"], "sc": r["_sch"], "cl": r["_cls"], "o": r["_out"], "n": int(r["_att"])}
                      for _, r in _pr.sort_values("_dt", ascending=False).iterrows()]
        PH["dates"] = _pr["_dt"].dropna().dt.strftime("%b %d").tolist()
        PH["by_day"] = _pr.assign(_d=_pr["_dt"].dt.date).groupby("_d").size().to_dict()
        PH["by_enum"] = _pr.groupby("_en").agg(calls=("_hh", "size"), parents=("_hh", "nunique"), ok=("_st", lambda x: int((x == 1).sum()))).reset_index().to_dict("records")
except Exception as e:
    say(f"  ! phone detail skipped: {e!r}")
ph_disp_n, n_ph_nr, n_ph_refused, n_ph_notreached, n_ph_agreed, n_ph_retry, n_ph_final, ph_school = {}, 0, 0, 0, 0, 0, 0, {}
N_ATT_ALL = N_ATT + n_ph_nr                                     # field visits only (n_ph_nr is 0 here)
conv_all = pct(N_C, N_ATT_ALL)
n_covered = N_C + n_ph_nr                                       # parents interviewed + parents tried by phone without success
remaining = max(N_TARGET - n_covered, 0)                        # households still open: interviewed, or already tried by phone without success, are taken off
COV_EXTRA = sum(v["actual"] - v["complete"] for v in COV_SCHOOL.values())    # children of closed-out schools reached but not interviewed (refused, wrong number, phone off ...)
if COV_SCHOOL:                                                  # target is counted in children (class lists), so take off the children interviewed + the ones reached without an interview
    remaining = max(N_TARGET - len(KIDC) - COV_EXTRA - n_ph_nr, 0)
days_left = math.ceil(remaining / per_day) if per_day and remaining else 0
proj_finish = next_workdays(last_ts, days_left)[-1] if (n_field_days and days_left) else last_ts

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
    {"l": "All attempts (field visits + phone calls)", "n": N_ATT_ALL},
    {"l": "Reached and consented", "n": n_consent},
    {"l": "Completed interviews", "n": int((D["_status"] == 1).sum())},
    {"l": "Counted in analysis (unique, in frame)", "n": N_C},
    {"l": "Bid recorded (WTP exercise)", "n": n_bid},
]

# ---- disposition
disp_codes = CH.get("status_survey", [])
_dn = {}
for v, l in disp_codes:
    _n = int((D["_status"] == v).sum())
    if _n:
        _dn[l] = _dn.get(l, 0) + _n
for l, n in ph_disp_n.items():                                  # phone outcomes join the same list (same label = same bar)
    _hit = next((k for k in _dn if k.replace("⁠", "").strip().lower() == l.lower()), None)
    _dn[_hit or l] = _dn.get(_hit or l, 0) + n
disp = sorted(({"l": l, "n": n} for l, n in _dn.items()), key=lambda x: -x["n"])
for x in disp:
    x["p"] = pct(x["n"], N_ATT_ALL)
    x["lab"] = f"{x['n']:,} · {x['p']:.0f}%"

# ---- assignment blocks (arm x order)
blocks = []
ARM_LAB = {0: "Control", 1: "Video 1", 2: "Video 2"}
ORD_LAB = {1: "beliefs asked before bid", 2: "beliefs asked after bid"}
for arm in (0, 1, 2):
    for o in (1, 2):
        tgt = int(((frame["ap_lm_arm"] == arm) & (frame["order_key"] == o)).sum()) if (frame is not None and not FRAME_IS_POOL) else 0
        if FRAME_IS_POOL:
            tgt = int(round(N_TARGET / 6))
        dn = int(((C["_arm"] == arm) & (C["_ord"] == o)).sum())
        if tgt == 0:
            tgt = max(dn, 1)
        p_ = pct(dn, tgt, 0)
        blocks.append({"bin": f"{ARM_LAB[arm]} · order {o}", "sub": ORD_LAB[o], "done": dn, "target": tgt, "pct": p_,
                       "status": "Completed" if dn >= tgt else ("In Progress" if dn else "Not Started"), "arm": arm})

# ---- classrooms: the class rule is applied class by class
def _binom_tail(n, p, k):
    """P(X >= k) for X ~ Binomial(n, p)"""
    if k <= 0:
        return 1.0
    if k > n:
        return 0.0
    p = min(max(float(p), 0.0), 1.0)
    return float(sum(math.comb(n, i) * p ** i * (1 - p) ** (n - i) for i in range(k, n + 1)))


_overall_p = float(np.nanmean(KIDC["_pclear"])) if KIDC["_pclear"].notna().any() else 0.0
CLS_COLOR = {"Secured": 6, "On track": 3, "At risk": 8, "Cannot reach": 2, "Not started": "muted"}
CLS_TONE = {"Secured": "good", "On track": "info", "At risk": "amber", "Cannot reach": "bad", "Not started": "grey"}
# final outcome of every household: Completed / Refused / Ineligible are final; Partial / Not reached / Not visited are still pending
_ST_LAB = {1: "Completed", 2: "Partial", 3: "Refused", 4: "Refused", 5: "Refused", 6: "Refused", 7: "Not reached", 8: "Not reached", 9: "Ineligible", 10: "Not reached"}
_hh_state = {}
for _h, _g in D.sort_values("_dt").groupby("_hh"):
    _last = _g["_status"].iloc[-1]
    _hh_state[int(_h)] = "Completed" if (_g["_status"] == 1).any() else _ST_LAB.get(int(_last) if pd.notna(_last) else 0, "Other")
_PENDING = ("Partial", "Not reached", "Not visited", "Other")
_hh_class = D.dropna(subset=["_ckey"]).groupby("_hh")["_ckey"].last().to_dict()          # class each household was sampled from
class_rows = []
for _ck in UNIVERSE:
    _g = KIDC[KIDC["_ckey"] == _ck]                      # one record per child interviewed about in this class
    _nd = len(_g)
    _sizes = _g["_tpar"].dropna()
    _size = _class_size(_ck)
    _cov = COV_CLASS.get((_short(_ck.split("|")[0]), int(_ck.split("|")[1])))
    if _cov:
        _size = _cov["actual"]
    _target = _size
    _thr = SCHOOL_THR_BY_NAME.get(_short(_ck.split("|")[0]), CLASS_THRESHOLD)
    _needed = int(_size * _thr)                          # same rule as the form: int(total_parents x threshold)
    _hits = int(np.nansum(_g["_hit"])) if _nd else 0
    _hhs = [h for h, c in _hh_class.items() if c == _ck]
    _states = [_hh_state.get(int(h), "Not visited") for h in _hhs]
    _refused = sum(1 for x in _states if x in ("Refused", "Ineligible"))
    if _cov:
        _rem = max(_size - _cov["complete"] - _cov["final"], 0)
        _refused = _cov["final"]
    elif len(_sizes) or HAS_FRAME_CLASS:
        _rem = max(_size - _nd - _refused, 0)            # parents of the class with no final outcome yet (refusals can no longer contribute)
    else:
        _rem = sum(1 for x in _states if x in _PENDING)
    _pbar = float(np.nanmean(_g["_pclear"])) if _nd >= 5 and _g["_pclear"].notna().any() else _overall_p
    _chance = _binom_tail(_rem, _pbar, _needed - _hits)
    if _nd == 0:
        _st = "Not started"
    elif _hits >= _needed:
        _st = "Secured"
    elif _hits + _rem < _needed:
        _st = "Cannot reach"
    elif _chance >= 0.5:
        _st = "On track"
    else:
        _st = "At risk"
    _sch, _gr, _sec = _ck.split("|")
    class_rows.append({"key": _ck, "label": _cls_label(_ck), "school": _short(_sch), "grade": int(_gr), "sec": _sec, "target": int(_target), "done": _cov["complete"] if _cov else _nd,
                       "size": _size, "needed": _needed, "thr": int(round(100 * _thr)), "hits": _hits, "rem": _rem, "refused": _refused, "rate": 100.0 * _hits / _size if _size else 0.0,
                       "chance": 1.0 if _st == "Secured" else (0.0 if _st == "Cannot reach" else _chance), "status": _st,
                       "incons": int(_sizes.nunique() > 1) if len(_sizes) else 0, "mean_bid": avg(_g["_bid"])})
class_rows.sort(key=lambda r: (r["school"], r["grade"], r["sec"]))
CLS_LABEL = {r["key"]: r["label"] for r in class_rows}
for _r in class_rows:
    if _r["done"] or _r["hits"]:
        say(f"  class: {_r['label']}: size {_r['size']} · rule {_r['thr']}% -> need {_r['needed']} · contributing {_r['hits']} · interviewed {_r['done']} · {_r['status']}")
n_cls = len(class_rows)
cls_n = {k: sum(1 for r in class_rows if r["status"] == k) for k in CLS_COLOR}
cls_expected = sum(r["chance"] for r in class_rows)
cls_schools = sorted({r["school"] for r in class_rows})

# ---- household coverage by area
wtp_all = C["_bid"]
wtp_mean = avg(wtp_all)
share_pos = sh(wtp_all, lambda s: s > 0)

school_blocks = []
_ph_by_key = {_nk(k): (k, v) for k, v in ph_school.items() if k}
SCHOOL_TARGET = {_short(SCHOOL_LABEL[int(_c)]): int(_n) for _c, _n in CFG.get("school_targets", {}).items() if int(_c) in SCHOOL_LABEL}   # parents per school, from the school list
for _lab_ in LIST_ROWS:                                       # schools with typed lists: the target is the list itself
    SCHOOL_TARGET[_short(_lab_)] = len(LIST_ROWS[_lab_])
for _sn in cls_schools + sorted(_short(l) for l in SCHOOL_LABEL.values() if _short(l) not in cls_schools):                  # every school in the form shows, started or not
    _cr = [r for r in class_rows if r["school"] == _sn]
    _t = SCHOOL_TARGET.get(_sn) or sum(r["target"] for r in _cr)
    _d = int(sum(1 for _k in KIDC["_ckey"].dropna() if _short(str(_k).split("|")[0]) == _sn))   # completed = children interviewed about (one per child), not parents
    _par = int((C["_sch"].map(_short) == _sn).sum())             # unique parents interviewed in this school (one survey per parent)
    _ph = _ph_by_key.pop(_nk(_sn), (None, {"nr": 0, "retry": 0, "final": 0}))[1]
    _reach = _d + _ph["nr"]
    _cv = COV_SCHOOL.get(_sn)
    if _cv:                                                      # list-checked sheet: actual parents, complete, and every other outcome
        _t, _d = _cv["actual"], _cv["complete"]
        _reach = _cv["actual"]
        _ph = {"nr": _reach - _d, "retry": _cv["off"] + _cv["notpicked"], "final": _cv["final"]}
    school_blocks.append({"bin": _sn, "sub": f"{len(_cr)} classes · {sum(1 for r in _cr if r['status'] == 'Secured')} secured" if _cr else "Not started yet",
                          "done": _d, "target": max(_t, 1), "pct": pct(_d, max(_t, 1), 0), "status": "Completed" if _d >= _t > 0 else ("In Progress" if _d else "Not Started"),
                          "att": _ph["nr"], "retry": _ph["retry"], "final": _ph["final"], "reach": _reach, "reach_pct": pct(_reach, max(_t, 1), 0), "par": _par,
                          **({"detail": " · ".join(f"{_cv[k]} {l}" for k, l in (("refused", "refused"), ("off", "phone off"), ("notpicked", "not picked"), ("wrong", "wrong number"), ("invalid", "invalid number"), ("sibling", "via sibling")) if _cv[k])} if _cv else {})})
for _k, (_sn, _ph) in _ph_by_key.items():                      # schools with phone calls but no field interview yet (class sizes not known: no target)
    school_blocks.append({"bin": _sn, "sub": "Phone calls only so far", "done": 0, "target": None, "pct": 0, "status": "In Progress",
                          "att": _ph["nr"], "retry": _ph["retry"], "final": _ph["final"], "reach": _ph["nr"], "reach_pct": None})

_nosch = int(C["_sch"].isna().sum())                          # completed interviews from earlier rounds with no school / class on record
if _nosch:
    school_blocks.append({"bin": "No school recorded", "sub": "Earlier rounds: school and class were not captured", "done": _nosch, "target": None, "pct": 0, "status": "Completed",
                          "att": 0, "retry": 0, "final": 0, "reach": _nosch, "reach_pct": None})
say("School progress: " + " | ".join(f"{b['bin']} {b['done']}/{b['target']}" for b in school_blocks))
n_cls_done = sum(1 for r in class_rows if r["rem"] == 0)                       # every parent has a final outcome
n_cls_80 = sum(1 for r in class_rows if 0 < r["rem"] <= 5)
ov_kpis = [
    K("✅", fmt_n(N_C), "Completed Interviews · Overall", f"Parents interviewed, all schools together · {META['pct_complete']:.0f}% of {N_TARGET:,} target · {len(KIDC):,} children covered (a parent with children in two classes counts in each class)", "up", "green"),
    *[K("🏫", fmt_n(b["done"]), f"Completed · {b['bin']}", f"{b['par']} unique parents · children counted once per child, so a parent with 2-3 children is counted 2-3 times · {b['pct']:.0f}% of {b['target']:,} children", "up", "teal") for b in school_blocks if b["done"] and b["target"]],
    K("🎯", fmt_n(remaining), "Households Remaining",
      (f"{N_TARGET:,} children in the classes − {len(KIDC):,} interviewed − {COV_EXTRA} reached but not interviewed ({', '.join(COV_SCHOOL)}: refused, wrong number, phone off, not picked up) = {remaining:,} still to reach" if COV_SCHOOL else
       f"{N_TARGET:,} target − {N_C} interviewed − {n_ph_nr} tried by phone without success ({n_ph_retry} to call again, {n_ph_final} refused / wrong number)" if n_ph_nr
       else f"{arm_target[0]} control · {arm_target[1]} video 1 · {arm_target[2]} video 2 targeted"), "navy", "navy"),
    K("🏁", f"{n_cls_done} / {n_cls}", "Classes Finished", f"Every parent interviewed or has a final outcome · {n_cls_80} more with 5 or fewer left", "up", "teal"),
    K("🏫", f"{cls_n['Secured']} / {n_cls}", "Classes With a Purifier Secured", f"{PCTM_SHORT} of a class contributing · {cls_n['On track']} more on track", "up", "purple"),
    K("🚪", fmt_n(N_ATT_ALL), "Total Attempts · Field + Phone", f"{N_C} completed ({conv_all:.0f}%) · {n_ph_nr} phone parents not interviewed" if n_ph_nr else f"{conv:.0f}% ended in a completed interview", "neutral", "amber"),
    *([K("📞", f"{n_covered} / {N_TARGET}", "Parents Covered So Far", f"{N_C} interviewed + {n_ph_nr} tried by phone without success · {pct(n_covered, N_TARGET, 0):.0f}% of {N_TARGET:,}", "neutral", "navy")] if n_ph_nr else []),
    K("💰", fmt_rs(wtp_mean), "Mean Willingness to Pay", f"{fmt_p(share_pos)} bid above zero", "up", ""),
    K("❓", n_dk_q_flagged, "Questions With High Don't-Know / Refused",
      f"{n_dk_q_flagged} of {n_dk_q_total} asked questions over {DK_Q_THRESHOLD:.0f}% Don't Know/Refused" if N_C else "No completed interviews yet",
      "down" if n_dk_q_flagged else "up", "amber" if n_dk_q_flagged else "green"),
]
ov_callouts = [
    {"cls": "navy", "icon": "⚪", "h": "Control — no video", "p": f"{arm_done[0]} of {arm_target[0]} completed ({pct(arm_done[0], arm_target[0], 0):.0f}%)."},
    {"cls": "teal", "icon": "🎬", "h": "Video 1 — information film", "p": f"{arm_done[1]} of {arm_target[1]} completed ({pct(arm_done[1], arm_target[1], 0):.0f}%)."},
    {"cls": "purple", "icon": "🎞️", "h": "Video 2 — information film", "p": f"{arm_done[2]} of {arm_target[2]} completed ({pct(arm_done[2], arm_target[2], 0):.0f}%)."},
    {"cls": "amber", "icon": "📉", "h": "Non-response", "p": f"{n_refused + n_ph_refused} refusals and {n_notreached + n_ph_notreached} not reached (phone off, no answer, wrong number, no one home) out of {N_ATT_ALL} attempts, field visits and phone calls together ({pct(n_refused + n_ph_refused + n_notreached + n_ph_notreached, N_ATT_ALL, 0):.0f}%)."},
    {"cls": "amber" if n_dk_q_flagged or n_dk_records else "teal", "icon": "❓", "h": "Don't know / non-response",
     "p": (f"{n_dk_q_flagged} of {n_dk_q_total} questions over {DK_Q_THRESHOLD:.0f}% Don't-Know/Refused/blank among respondents asked; "
           f"{n_dk_records} interview(s) with {DK_RECORD_MIN}+ such answers are listed to verify under Field Operations.") if N_C else "No completed interviews yet."},
]

# ---- key findings (auto-written from the data)
kf = []
if n_cls and N_C:
    kf.append({"cls": "teal", "html": f"<strong>Classrooms:</strong> <strong>{cls_n['Secured']}</strong> of {n_cls} classes have already reached the {PCTM} contribution rule and would get a purifier; "
               f"<strong>{cls_n['On track']}</strong> more are on track and <strong>{cls_n['At risk']}</strong> are at risk. Expected by the end of fieldwork: about <strong>{cls_expected:.0f}</strong> of {n_cls}."})
if N_C:
    kf.append({"cls": "primary", "html": f"<strong>Willingness to pay:</strong> the average parent would contribute <strong>{fmt_rs(wtp_mean)}</strong> "
               f"(median {fmt_rs(med(wtp_all))}); <strong>{fmt_p(share_pos)}</strong> bid above zero and "
               f"<strong>{fmt_p(sh(wtp_all, lambda s: s >= CAP))}</strong> bid the Rs&nbsp;{CAPS} maximum. "
               f"Whether a class gets a purifier depends on the share of <em>its own</em> parents contributing (see Classrooms)."})
    aq = avg(num(C, "outdoor_air_quality"))
    aqi_ok = f_multi("beleifs_air_pollute", "belief_air")(C)
    aqi_share = next((i["p"] for i in aqi_ok["items"] if i["l"].lower().startswith("levels")), None) if aqi_ok else None
    kf.append({"cls": "navy", "html": f"<strong>Awareness:</strong> parents rate outdoor air quality <strong>{r1(aq) if aq is not None else '—'}/10</strong> on the pollution scale, "
               f"yet only <strong>{fmt_p(aqi_share)}</strong> connect the AQI with levels of air pollution."})
arms_ok = all(int((C["_arm"] == v).sum()) >= 8 for v in (0, 1, 2)) if N_C else False
if arms_ok:
    m0, m1, m2 = (avg(C.loc[C["_arm"] == v, "_bid"]) for v in (0, 1, 2))
    kf.append({"cls": "purple", "html": f"<strong>First read on the videos:</strong> mean bid is <strong>{fmt_rs(m0)}</strong> in Control, <strong>{fmt_rs(m1)}</strong> after Video 1 "
               f"and <strong>{fmt_rs(m2)}</strong> after Video 2. Treat as indicative until fieldwork is complete."})
kf.append({"cls": "teal", "html": f"<strong>Pace:</strong> {per_day:.1f} completed interviews per field day; "
           f"<strong>{fmt_n(remaining)}</strong> households remain" + (f", so the target is reached around <strong>{fdate(proj_finish)}</strong> at this rate." if days_left else ".")})

overview = {
    "id": "overview", "tab": "📊 Overview", "eyebrow": "Section 01", "title": "Programme at a Glance",
    "blurb": f"Where fieldwork stands against the {N_TARGET:,}-household sample across {n_cls} classes in {len(cls_schools)} schools: completed interviews, how many classrooms are on course for a purifier, "
             "how the three randomised study groups are filling, and the headline findings so far.",
    "seg": False, "blocks": [
        {"t": "kpis", "d": ov_kpis},
        {"t": "callouts", "d": ov_callouts},
        {"t": "insights", "d": kf, "row": False},
        {"t": "grid", "cols": 1, "cards": [
            card("lines", "ovDaily", "Daily Completed Interviews", "Completed interviews per field day",
                 {"x": [x["d"] for x in daily], "series": [{"name": "Completed", "slot": 2, "y": _vals, "fill": True, "pts": True, "labels": True}]},
                 opt={"yfmt": "n"}, full=True)]},
        {"t": "grid", "cols": 3, "cards": [
            card("gauge", "ovProg", "Overall Progress vs Target", f"Completed interviews against the {N_TARGET:,} parents in the {n_cls} classes seen so far" if FRAME_IS_POOL else f"Completed interviews against the {N_TARGET:,}-household frame", {"done": N_C, "target": N_TARGET}),
            card("donut", "ovArm", "Completed by Study Arm", "Control · Video 1 · Video 2",
                 {"items": [{"l": ARM_LAB[v], "n": arm_done[v], "p": pct(arm_done[v], N_C)} for v in (0, 1, 2)], "base": N_C}, opt={"colors": [7, 3, 5], "fmt": "n"}),
            card("hbar", "ovDisp", "Attempt Disposition", "Outcome of every fieldwork attempt recorded so far", {"items": disp, "base": N_ATT_ALL}, var="status_survey", opt={"color": 2, "fmt": "n"})]},
        {"t": "grid", "cols": 2, "cards": [
            card("funnel", "ovFunnel", "From Attempt to Analysed Interview", "How many attempts survive each stage", {"rows": funnel}),
            card("blocks", "ovBlocks", "Progress by School", "Completed interviews against the parents listed in each school's classes", {"rows": school_blocks})]},
        {"t": "grid", "cols": 1, "cards": [
            card("table", "ovDkQ", "Don't Know / Refused / Not Sure, by Question",
                 f"Every question the instrument gives a Don't Know / Refused / Not-sure option to, and the share of respondents it was actually asked of (skip-pattern cases excluded) who answered Don't Know, Refused or Not sure. Blanks are not counted. Flagged above {DK_Q_THRESHOLD:.0f}%.",
                 dk_table, opt={"sortKey": "p"}, full=True)]},
        {"t": "note", "html": ((f"Sample: <strong>{n_cls}</strong> classes in <strong>{len(cls_schools)}</strong> schools; the parents of each class are interviewed, about <strong>{N_TARGET:,}</strong> interviews in all "
                                  f"(schools, classes and class sizes appear as the interviews report them, so this total grows as new classes start). ") if FRAME_IS_POOL else
                                 (f"Sampling frame: <strong>{N_TARGET:,}</strong> households in <strong>{n_cls}</strong> classes, pre-assigned to study arm and question order (Control = {arm_target[0]}, Video 1 = {arm_target[1]}, Video 2 = {arm_target[2]}). ")) +
                             "Completed = <code>status_survey</code> is “Completed”, counted once per household. Attempts add the phone calls: each parent is counted once however many times they were called, and parents who could not be interviewed by phone (switched off, did not pick up, wrong number, refused) count as attempts and non-response, not as completed interviews. Research Solutions (M&amp;A Research Solutions LLC) | www.rs.org.pk"},
    ]}
PANELS.append(overview)

# ====================================================================== PANEL - CLASSROOMS (the class rule)
_cls_sorted = sorted(class_rows, key=lambda r: (-(r["hits"] / r["needed"] if r["needed"] else 0), r["label"]))
_pill = lambda k: {"v": {"Secured": 4, "On track": 3, "At risk": 2, "Cannot reach": 1, "Not started": 0}[k], "t": k, "tone": CLS_TONE[k]}
cls_table = {"cols": [{"k": "school", "l": "School"}, {"k": "cls", "l": "Class"}, {"k": "rule", "l": "Rule", "num": True}, {"k": "size", "l": "Class size", "num": True},
                      {"k": "done", "l": "Interviewed", "num": True}, {"k": "pend", "l": "Still to do", "num": True}, {"k": "hits", "l": "Contributing", "num": True},
                      {"k": "need", "l": f"Needed ({PCTM_SHORT})", "num": True}, {"k": "rate", "l": "Share contributing", "num": True},
                      {"k": "chance", "l": "Chance of reaching the rule", "num": True}, {"k": "st", "l": "Status", "verdict": True}],
             "rows": [{"school": r["school"], "rule": {"v": r["thr"], "t": f"{r['thr']}%"}, "cls": f"Class {r['grade']}" + (f"-{r['sec']}" if r["sec"] else ""), "size": r["size"],
                       "done": {"v": r["done"], "t": f"{r['done']} / {r['target']}"}, "pend": r["rem"], "hits": r["hits"], "need": r["needed"],
                       "rate": {"v": round(r["rate"], 1), "t": f"{r['rate']:.0f}%"}, "chance": {"v": round(100 * r["chance"], 1), "t": f"{100 * r['chance']:.0f}%"},
                       "st": _pill(r["status"])} for r in class_rows],
             "read": f"<strong>{cls_n['Secured']}</strong> of {n_cls} classes are already at or above {PCTM}. A class can still change until every parent is interviewed."}
_prog = lambda r: 100.0 * r["hits"] / r["needed"] if r["needed"] else 0.0            # contributors so far as a share of the number needed
cls_bars = {"items": [{"l": r["label"], "v": round(_prog(r), 1), "p": round(_prog(r), 1), "n": r["hits"], "lab": f"{r['hits']} of {r['needed']}", "slot": CLS_COLOR[r["status"]],
                       "tip": f"{r['hits']} contributing, {r['needed']} needed ({r['thr']}% of {r['size']}) · {r['done']} of {r['target']} interviewed, {r['rem']} still to do · {r['status']}"} for r in _cls_sorted],
            "base": None}
if _cls_sorted:
    _near = [r for r in _cls_sorted if r["status"] in ("On track", "At risk")]
    _near.sort(key=lambda r: r["needed"] - r["hits"])
    cls_bars["read"] = (f"Furthest ahead: <strong>{_cls_sorted[0]['label']}</strong> ({_cls_sorted[0]['hits']} of {_cls_sorted[0]['needed']} needed). Furthest behind: <strong>{_cls_sorted[-1]['label']}</strong> ({_cls_sorted[-1]['hits']} of {_cls_sorted[-1]['needed']}). "
                        + (f"Closest to the line and not yet secured: <strong>{_near[0]['label']}</strong> ({_near[0]['hits']} of {_near[0]['needed']} needed)." if _near else ""))
cls_status = {"items": [{"l": k, "n": v, "p": pct(v, n_cls)} for k, v in cls_n.items() if v], "base": n_cls,
              "read": f"<strong>{cls_n['Secured']}</strong> secured, <strong>{cls_n['On track']}</strong> on track, <strong>{cls_n['At risk']}</strong> at risk, <strong>{cls_n['Cannot reach']}</strong> cannot reach {PCTM}, <strong>{cls_n['Not started']}</strong> not started."}
_rates = [_prog(r) for r in class_rows if r["done"] > 0]
_bins = [(0, 25), (25, 50), (50, 75), (75, 100), (100, 100000)]
_bl = ["Under 25%", "25–49%", "50–74%", "75–99%", "Secured (100%+)"]
cls_dist = {"items": [{"l": l, "n": sum(1 for x in _rates if lo <= x < hi), "p": pct(sum(1 for x in _rates if lo <= x < hi), len(_rates)),
                       "slot": 6 if lo >= 100 else (3 if lo >= 75 else (8 if lo >= 50 else 2))} for (lo, hi), l in zip(_bins, _bl)], "base": len(_rates), "hist": True}
if cls_dist["items"] and _rates:
    _top = max(cls_dist["items"], key=lambda i: i["n"])
    cls_dist["read"] = f"Most classes ({_top['n']} of {len(_rates)}) are in the <strong>{_top['l']}</strong> group."
_sch_rows = []
for _sn in cls_schools:
    _cr = [r for r in class_rows if r["school"] == _sn]
    _sch_rows.append({"l": _sn, "p": [pct(sum(1 for r in _cr if r["status"] == k), len(_cr)) for k in CLS_COLOR], "c": [sum(1 for r in _cr if r["status"] == k) for k in CLS_COLOR], "n": len(_cr)})
_all_ok = [r["l"] for r in _sch_rows if r["c"][0] + r["c"][1] == r["n"]]
cls_school = {"cats": list(CLS_COLOR), "rows": _sch_rows,
              "read": f"<strong>{len(_all_ok)}</strong> of {len(_sch_rows)} schools have every class secured or on track" + (f" ({', '.join(_all_ok[:4])}{'…' if len(_all_ok) > 4 else ''})." if _all_ok else ".")}
by_arm_hit = [(ARM_LAB[v], sh(C.loc[C["_arm"] == v, "_hit"], lambda x: x == 1), int((C["_arm"] == v).sum())) for v in (0, 1, 2)]
cls_arm = {"items": [{"l": f"{l} (n = {n})", "v": r1(x), "p": r1(x), "n": n, "slot": [7, 3, 5][i]} for i, (l, x, n) in enumerate(by_arm_hit) if x is not None], "base": N_C}
if len(cls_arm["items"]) == 3:
    cls_arm["read"] = ("Share whose random price cleared their bid: " + " · ".join(f"{a}: <strong>{x:.0f}%</strong>" for a, x, _ in by_arm_hit) +
                       f". More contributors per class means more classes reach the {PCTM} rule.")

PANELS.append({
    "id": "classes", "tab": "🏫 Classrooms", "eyebrow": f"Section 02 · The {PCTM_SHORT} rule, class by class", "title": "Which Classrooms Get a Purifier?",
    "blurb": f"A purifier goes into a classroom only if at least {PCTM} of that class's parents end up contributing. Each parent's contribution is decided by a random price against their own bid. Parents we have not reached count as not contributing, so a class can still move until every parent has a final outcome.",
    "seg": False, "blocks": [
        {"t": "insight_static", "cls": "navy", "html": "<strong>🏫 How a class gets its purifier:</strong> each parent names the most they would pay. A random price is drawn: if it is at or below the bid, that parent contributes the random price. "
                   f"If <strong>at least {PCTM} of all parents in the class</strong> (for a class of 28, that is {int(28 * CLASS_THRESHOLD)} parents) end up contributing, the whole class gets a purifier; otherwise nobody pays. Parents who cannot be reached do not count."},
        {"t": "kpis", "d": [
            K("🏫", n_cls, "Classes in the Study", f"{len(cls_schools)} schools · classes {int(min(r['grade'] for r in class_rows)) if class_rows else '—'}–{int(max(r['grade'] for r in class_rows)) if class_rows else '—'}", "navy", "navy"),
            K("✅", cls_n["Secured"], "Purifier Secured", f"{PCTM_SHORT} of the class already contributing", "up", "green"),
            K("📈", cls_n["On track"], "On Track", "Likely to get there as parents are interviewed", "up", "teal"),
            K("⚠️", cls_n["At risk"], "At Risk", "Unlikely at the current contribution rate", "neutral", "amber"),
            K("⛔", cls_n["Cannot reach"], f"Cannot Reach {PCTM_SHORT}", "Even if every remaining parent contributed", "down", ""),
            K("🎯", f"{cls_expected:.1f}", "Expected Classes With a Purifier", f"Of {n_cls}, at the end of fieldwork, from current bids", "neutral", "purple")]},
        {"t": "grid", "cols": 2, "cards": [
            card("donut", "clStatus", "Where do the classes stand?", "Number of classes in each status", cls_status, opt={"colors": [CLS_COLOR[k] for k in CLS_COLOR if cls_n[k]], "fmt": "n"}),
            card("bar", "clDist", "How many classes are near the target?", "Classes grouped by how much of the contributors they need they already have", cls_dist, opt={"fmt": "n"})]},
        {"t": "grid", "cols": 1, "cards": [
            card("hbar", "clBars", f"How close is each class to the {PCTM_SHORT} line?", f"Contributors so far as a share of the number the class needs ({PCTM} of the class). The dashed line is 100%: class secured.", cls_bars,
                 opt={"fmt": "num1", "sfx": "%", "ref": 100, "refLabel": "Number needed", "legend": [[k, CLS_COLOR[k]] for k in ("Secured", "On track", "At risk", "Cannot reach")]}, full=True)]},
        {"t": "grid", "cols": 2, "cards": [
            card("likert", "clSchool", "How do each school's classes split?", "Each bar is one school; the colours are its classes' status", cls_school, opt={"pal": "status", "counts": True, "slots": [CLS_COLOR[k] for k in CLS_COLOR]}),
            card("hbar", "clArm", "Do the videos raise the share who contribute?", "Share whose random price cleared their bid, by study group", cls_arm, opt={"fmt": "num1", "sfx": "%", "color": 1})]},
        {"t": "grid", "cols": 1, "cards": [
            card("table", "clTable", "Every class, one line each", f"Click a column to sort, or filter by school. “Chance” is the probability of reaching {PCTM} once the parents still to interview are counted.", cls_table,
                 opt={"sortKey": "rate", "filterKey": "school", "filterVals": ["All"] + cls_schools}, full=True)]},
        {"t": "note", "html": f"Needed = {PCTM} of the class size reported in the interviews (rounded down), the same rule the questionnaire uses"
                             + (f" (schools finished under the earlier protocol keep their own rule: {', '.join(f'{n} {int(round(100 * t))}%' for n, t in SCHOOL_THR_BY_NAME.items())})" if SCHOOL_THR_BY_NAME else "") + ". Contributing counts parents whose random price was at or below their bid "
                             f"(the revised bid and new draw where a parent changed their bid). Chance of reaching {PCTM} assumes parents still to be interviewed contribute at the class's average rate so far."},
    ]})


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
if HAS_PLAN:
    qa("Class not matched to the school plan", "Child records whose school, class or section is not one of the planned classes (not counted in any class)", N_UNMATCHED, len(KIDC))
qa("Class size entered inconsistently", "Classes where interviews report different class sizes", sum(r["incons"] for r in class_rows), n_cls)
qa("Many don't-know / non-response answers",
   f"Completed interviews where {DK_RECORD_MIN} or more of the {n_dk_q_total} don't-know-eligible questions came back Don't Know, Refused, or Not sure",
   n_dk_records, N_C)
n_flag_total = sum(r["n"] for r in qa_rows)

# ---- enumerator scorecard
enum_rows = []
overall_cap = sh(C["_bid"], lambda s: s >= CAP) or 0
for code, grp in D.groupby("_enum", dropna=True):
    lab = enum_label(code)
    cg = C[C["_enum"] == code]
    n_att, n_cmp = len(grp), int((grp["_status"] == 1).sum())
    cap = sh(cg["_bid"], lambda s: s >= CAP)
    flags = []
    if len(cg) and int((cg["_min"] < MIN_DUR).sum()):
        flags.append(f"{int((cg['_min'] < MIN_DUR).sum())} short")
    if len(cg) >= 5 and cap is not None and cap >= 40 and cap >= 2 * max(overall_cap, 1):
        flags.append(f"bids cluster at Rs {CAPS}")
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
        {"t": "grid", "cols": 1, "cards": [
            card("table", "opsDkQ", "Don't Know / Refused / Not Sure, by Question",
                 f"Every question the instrument gives a Don't Know / Refused / Not-sure option to, and the share of respondents it was actually asked of who answered Don't Know, Refused or Not sure. Blanks are not counted. Flag line: {DK_Q_THRESHOLD:.0f}%.",
                 dk_table, opt={"sortKey": "p"}, full=True)]},
        {"t": "grid", "cols": 1, "cards": [card("table", "opsEnum", "Enumerator Scorecard", "Attempts, conversion, interview length and bid pattern by field team member. Click a column to sort.",
                                                {"cols": [{"k": "e", "l": "Enumerator"}, {"k": "att", "l": "Attempts", "num": True}, {"k": "cmp", "l": "Completed", "num": True},
                                                          {"k": "conv", "l": "Conversion", "num": True}, {"k": "med", "l": "Median min", "num": True},
                                                          {"k": "mean", "l": "Mean bid", "num": True}, {"k": "cap", "l": f"Bids at Rs {CAPS}", "num": True},
                                                          {"k": "flags", "l": "To verify", "flags": True}], "rows": enum_rows}, full=True, opt={"sortKey": "cmp"})]},
        {"t": "grid", "cols": 3, "cards": [
            card("bar", "opsDur", "Interview Duration", "Completed interviews by elapsed time: a tight distribution is the expected shape", dur_hist, opt={"color": 4, "fmt": "pct"}),
            card("bar", "opsHour", "Interview Start Hour", "Completed interviews by the hour of day they began (24-hour clock)", hour_hist, opt={"color": 1, "fmt": "n"}),
            card("hbar", "opsEnum2", "Output by Enumerator", "Completed interviews per field team member", enum_bar, opt={"color": 1, "fmt": "n"})]},
        {"t": "grid", "cols": 1, "cards": [card("recon", "opsRecon", "Data Source Check", "How the CSV and Stata exports compare before they are combined", recon_card)]},
        {"t": "insights", "d": ops_ins, "row": True},
        {"t": "note", "html": "Enumerators are shown by staff code. Duration is start to submission. Checks are prompts for verification, never findings of misconduct. Don't-Know/Refused shares are computed only among respondents each question was actually asked of (skip-pattern cases excluded). The projected finish extrapolates the average daily completion rate and skips Sundays."},
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


_gvals = sorted(int(x) for x in C["_gradeN"].dropna().unique()) or [1]
_GR_EDGES = _gvals + [_gvals[-1] + 1]
_GR_LABELS = [f"Class {g}" for g in _gvals]
PANELS.append({
    "id": "child", "tab": "🎒 Child & School", "eyebrow": "Section 04 · School, commute & child health", "title": "Child, School & Health",
    "blurb": "The child at the centre of the survey: grade and recent exam results, how they get to school, how much schooling illness has cost them, and whether parents connect that illness to air pollution.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(child_kpis)},
        {"t": "grid", "cols": 3, "cards": [
            card("bar", "chGrade", "Child's grade", "Class of the child in the study classroom", SD(f_hist(lambda d: d["_grade"], _GR_EDGES, _GR_LABELS)), var="grade_child", opt={"color": 1}),
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
BIN_E = [0, 1, 251, 501, 751, 1001, 1501, 2001, CAP, CAP + 1]
BIN_L = ["Rs 0", "1–250", "251–500", "501–750", "751–1,000", "1,001–1,500", "1,501–2,000", f"2,001–{CAP - 1:,}", f"Rs {CAPS}"]
BIN_E[0] = -1


def wtp_kpis(df):
    b = df["_bid"]
    return [
        K("💰", fmt_rs(avg(b)), "Mean Willingness to Pay", f"Median {fmt_rs(med(b))}", "up", "green"),
        K("🙋", fmt_p(sh(b, lambda s: s > 0)), "Bid Above Zero", f"{fmt_p(sh(b, lambda s: s >= CAP))} bid the Rs {CAPS} maximum", "neutral", "teal"),
        K("🎯", fmt_p(sh(b, lambda s: s >= 500)), "Bid Rs 500 or More", f"{fmt_p(sh(b, lambda s: s >= 1000))} bid Rs 1,000 or more", "up", "purple"),
        K("🔄", fmt_p(sh(num(df, "change_pay"), lambda s: s == 1)), "Revised Their Bid", f"Mean bid {fmt_rs(avg(df['_bid0']))} → {fmt_rs(avg(b))}", "neutral", "amber"),
        K("🎲", fmt_p(pct(int(df["_hit"].sum()), int(df["_hit"].notna().sum())) if df["_hit"].notna().any() else None), "Would Have Paid", "Bid was at or above the random price drawn", "navy", "navy"),
        K("💸", fmt_rs(avg(df["_paid"])), "Mean Amount Paid", "Realised by the random-price rule", "navy", ""),
    ]


def others_belief(df):
    tot = pd.Series(np.nan, index=df.index)
    for c in ("total_parents", "total_parents_2"):
        tot = tot.fillna(num(df, c))
    rem = tot - 1
    parts = [("Not surveyed", "other_parents_fail"), ("Rs 0", "other_parents_0"), ("Rs 1–500", "other_parents_500"), ("Rs 501–1,000", "other_parents_1000"),
             ("Rs 1,001–1,500", "other_parents_1500")]
    if "other_parents_2499" in FORM_VARS:                                   # current instrument: ranges run up to the Rs 2,500 cap
        parts += [("Rs 1,501–2,000", "other_parents_2000"), ("Rs 2,001–2,499", "other_parents_2499"), (f"Rs {CAPS}", "other_parents_2500")]
    else:                                                                   # pilot instrument
        parts += [("Rs 1,501–1,999", "other_parents_1999"), ("Rs 2,000", "other_parents_2000")]
    its = []
    for lab, c in parts:
        v = num(df, c).fillna(num(df, c + "_2")) / rem
        v = v[rem > 0].dropna()
        if len(v):
            its.append({"l": lab, "v": r1(100 * float(v.mean())), "n": len(v)})
    return {"items": its, "base": int(rem.notna().sum())} if its else None


def _expected_others(df):
    """how many of the other parents in the class the respondent expects to contribute: their own revised
    guess where they changed it, else the figure worked out from their bid-range answers (pilot: other_parents)"""
    out = pd.Series(np.nan, index=df.index)
    for sfx in ("", "_2"):                                                  # question order 1 / 2 use different copies of the block
        est = num(df, "estimated_contributors" + sfx)
        rev = num(df, "other_parents_revised" + sfx)
        chk = num(df, "estimate_check" + sfx)
        out = out.fillna(rev.where(chk.isin([2, 3]), est))
    return out.fillna(num(df, "other_parents"))


def wtp_ins(df):
    b = df["_bid"]
    exp_ = None
    tot = pd.Series(np.nan, index=df.index)
    for c in ("total_parents", "total_parents_2"):
        tot = tot.fillna(num(df, c))
    op = _expected_others(df)
    ratio = (op / (tot - 1)).replace([np.inf, -np.inf], np.nan).dropna()
    if len(ratio):
        exp_ = 100 * float(ratio.mean())
    real = pct(int(df["_hit"].sum()), int(df["_hit"].notna().sum())) if df["_hit"].notna().any() else None
    out = [{"cls": "primary", "html": f"<strong>How much:</strong> the average bid is <strong>{fmt_rs(avg(b))}</strong> (median {fmt_rs(med(b))}); <strong>{fmt_p(sh(b, lambda s: s >= 500))}</strong> of parents would pay Rs 500 or more. Whether a class gets its purifier depends on its own parents: see the Classrooms tab."}]
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
        add("Study group", ARM_LAB[v], C["_arm"] == v, [7, 3, 5][v])
    out = []
    for g, l, ser, slot in rows:
        ser = ser.dropna()
        if len(ser) >= MIN_CELL:
            out.append({"g": g, "l": l, "m": r1(float(ser.mean()), 0), "n": int(len(ser)), "slot": slot})
    return {"rows": out, "unit": "Rs", "overall": r1(avg(C["_bid"]), 0)} if out else None


PRICE_STEPS = [100, 250, 500, 1000, 1500, 2000, CAP] if CAP > 2000 else [100, 250, 500, 1000, 1500, 2000]


def price_bars(df):
    b = df["_bid"].dropna()
    if not len(b):
        return None
    return {"items": [{"l": f"Rs {p:,} or more", "n": int((b >= p).sum()), "p": pct(int((b >= p).sum()), len(b))} for p in PRICE_STEPS], "base": int(len(b))}


PANELS.append({
    "id": "wtp", "tab": "💰 Willingness to Pay", "eyebrow": "Section 06 · Incentive-compatible elicitation (BDM)", "title": "Willingness to Pay for a Classroom Air Purifier",
    "blurb": f"Each parent states the most they would contribute towards an air purifier for their child's classroom, then a random price is drawn: they pay the random price only if their bid is at least as high, which makes an honest bid the best strategy. A classroom gets a purifier only if {PCTM} of its own parents end up contributing (see the Classrooms tab).",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(wtp_kpis)},
        {"t": "grid", "cols": 1, "cards": [
            card("bar", "wtDist", "Distribution of final bids", "Amount each parent was willing to contribute after any revision, in rupees", SD(f_hist(lambda d: d["_bid"], BIN_E, BIN_L)), var="contribute_will / change_pay_w", opt={"color": 6, "fmt": "pct"}, full=True)]},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "wtDemand", "How many parents would pay at least this much?", "Share of parents whose bid is at or above each amount", SD(price_bars), opt={"color": 6, "fmt": "pct"}),
            card("hbar", "wtGroup", "Average amount parents would pay, by group", "Average final bid for each group; not affected by the filter above", wtp_groups(), opt={"grouped": True}, tall=True, auto=True)]},
        {"t": "grid", "cols": 3, "cards": [
            card("hbar", "wtOthers", "What parents expect other parents to do", "Average expected split of the other parents in the class, % of class", SD(others_belief), opt={"color": 5, "fmt": "num1", "sfx": "%"}),
            card("donut", "wtChange", "Revised their bid after seeing the result", "Would you like to change how much you are willing to contribute?", SD(f_dist("change_pay", "yesno")), var="change_pay", opt={"colors": [4, "muted"]}),
            card("donut", "wtCertain", "Certainty about the bid", "How certain parents are of their choice", SD(f_dist("certain_choice", "choice", drop_zero=True)), var="certain_choice", opt={"colors": [6, 3, 1, 4, "muted"]})]},
        {"t": "grid", "cols": 2, "cards": [
            card("hbar", "wtWhy", "Why not pay more", f"Reason given by parents who bid below the Rs {CAPS} maximum and kept their bid", SD(f_dist("more_pay", "reason", drop_zero=True)), var="more_pay", opt={"color": 2}),
            card("bar", "wtPractice", "Practice round bids", "Rs 0–20 practice round before the real exercise", SD(f_hist(lambda d: num(d, "contribute_will_practice"), [-1, 1, 6, 11, 16, 21], ["Rs 0", "1–5", "6–10", "11–15", "16–20"])), var="contribute_will_practice", opt={"color": 1})]},
        {"t": "grid", "cols": 2, "cards": [
            card("donut", "wtSel", "If selected and the price is above the bid", "Comprehension: is the Rs 1,000 bonus still received?", SD(f_dist("selected_parent", "yesno")), var="selected_parent", opt={"colors": [6, 2]}),
            card("donut", "wtNot", "If not selected: which bid is used", "Comprehension of the second-stage rule", SD(f_dist("not_selected_parent", "bid_use")), var="not_selected_parent", opt={"colors": [6, 2, 4]})]},
        {"t": "insights", "d": SD(wtp_ins), "row": True},
        {"t": "note", "html": "Final bid = the revised bid where a parent changed it, otherwise the original. “Would have paid” compares the final bid with the final random price drawn in the interview."},
    ]})

# ====================================================================== PANEL - VIDEO EFFECTS
arm_masks = {v: C["_arm"] == v for v in (0, 1, 2)}
ARM_SLOT = {0: 7, 1: 3, 2: 5}
n_arm = {v: int(arm_masks[v].sum()) for v in (0, 1, 2)}
m_arm = {v: avg(C.loc[arm_masks[v], "_bid"]) for v in (0, 1, 2)}
te_kpis = []
for v in (0, 1, 2):
    m = m_arm[v]
    dlt = None if v == 0 or m_arm[0] is None or m is None else m - m_arm[0]
    te_kpis.append(K(["⚪", "🎬", "🎞️"][v], fmt_rs(m), f"{ARM_LAB[v]}: Average Bid", f"n = {n_arm[v]}" + ("" if dlt is None else f" · {dlt:+,.0f} vs Control"), "up" if (dlt or 0) > 0 else "neutral", ["navy", "teal", "purple"][v]))
ords = {v: C["_ord"] == v for v in (1, 2)}
_ob = {o: avg(C.loc[ords[o], "_bid"]) for o in (1, 2)}
te_kpis.append(K("🔀", fmt_rs(_ob[2]), "Asked About Others After Own Bid", "" if (_ob[1] is None or _ob[2] is None) else f"{_ob[2] - _ob[1]:+,.0f} vs asked before", "neutral", "amber"))

OUTCOMES = [
    ("Average bid (Rs)", "rs", lambda d: d["_bid"]),
    ("Average bonus-round bid (Rs)", "rs", lambda d: d["_bonus"]),
    ("Would have paid (random price cleared the bid)", "pp", lambda d: d["_hit"]),
    ("Bid above zero", "pp", lambda d: (d["_bid"] > 0).astype(float).where(d["_bid"].notna())),
    (f"Bid at the Rs {CAPS} maximum", "pp", lambda d: (d["_bid"] >= CAP).astype(float).where(d["_bid"].notna())),
    ("See ≥50% higher child risk", "pp", lambda d: num(d, "risk_air_child").where(num(d, "risk_air_child").isin([1, 2, 3, 4, 5])).map(lambda v: np.nan if pd.isna(v) else float(v >= 4))),
    ("Expect ≥10% exam gain from a purifier", "pp", lambda d: num(d, "score_exams_ap").map(lambda v: np.nan if pd.isna(v) else float(v >= 4))),
    ("Support publicly funded purifiers", "pp", lambda d: num(d, "support_policy").map(lambda v: np.nan if pd.isna(v) else float(v in (1, 2)))),
    ("Support a small mandatory school fee", "pp", lambda d: num(d, "small_fee").map(lambda v: np.nan if pd.isna(v) else float(v == 1))),
    ("Say air pollution should be a top school priority", "pp", lambda d: num(d, "top_prio").map(lambda v: np.nan if pd.isna(v) else float(v == 1))),
]
eff_rows, wins = [], {0: 0, 1: 0, 2: 0}
for oname, unit, getter in OUTCOMES:
    ser = getter(C); k_ = 1 if unit == "rs" else 100
    vals = [float(ser[arm_masks[v]].mean()) * k_ if ser[arm_masks[v]].notna().any() else None for v in (0, 1, 2)]
    ok = [x for x in vals if x is not None]
    best = vals.index(max(ok)) if len(ok) >= 2 and max(ok) > min(ok) else None
    if best is not None:
        wins[best] += 1
    fmt_o = lambda x: "—" if x is None else (f"Rs {x:,.0f}" if unit == "rs" else f"{x:.0f}%")
    eff_rows.append({"o": oname, **{key: {"v": vals[i], "t": fmt_o(vals[i]), "best": best == i} for i, key in enumerate(("c", "v1", "v2"))}})
eff_data = ({"cols": [{"k": "o", "l": "Question"}, {"k": "c", "l": "Control", "num": True, "hl": True}, {"k": "v1", "l": "Video 1", "num": True, "hl": True}, {"k": "v2", "l": "Video 2", "num": True, "hl": True}],
             "rows": eff_rows, "read": f"The highest group is Video 1 on <strong>{wins[1]}</strong> of {len(eff_rows)} questions, Video 2 on <strong>{wins[2]}</strong> and Control on <strong>{wins[0]}</strong>."}
            if eff_rows and N_C >= 3 else None)

BAL = [("Age (average, years)", "num", lambda d: d["_age"]),
       ("Monthly income (median, Rs)", "med", lambda d: d["_income"]),
       ("Household size (average)", "num", lambda d: d["_hhsize"]),
       ("Female respondent", "sh", lambda d: (num(d, "gender") == 2).astype(float).where(num(d, "gender").notna())),
       ("Urban area", "sh", lambda d: (num(d, "area_class") == 2).astype(float).where(num(d, "area_class").notna())),
       ("Owns an air conditioner", "sh", lambda d: (num(d, "asset_ac") == 1).astype(float).where(num(d, "asset_ac").notna())),
       ("Respiratory illness at home", "sh", lambda d: (num(d, "hh_resp_ill") == 1).astype(float).where(num(d, "hh_resp_ill").notna())),
       ("Outdoor pollution rating (average, 0–10)", "num", lambda d: num(d, "outdoor_air_quality")),
       ("Child's class (average)", "num", lambda d: d["_grade"])]
bal_rows = []
for name, kind, getter in BAL:
    ser = getter(C)
    groups = [ser[arm_masks[v]] for v in (0, 1, 2)]
    if kind == "sh":
        vals = [sh(g, lambda x: x == 1) for g in groups]; t = lambda x: "—" if x is None else f"{x:.0f}%"
    elif kind == "med":
        vals = [med(g) for g in groups]; t = lambda x: "—" if x is None else f"{x:,.0f}"
    else:
        vals = [avg(g) for g in groups]; t = lambda x: "—" if x is None else f"{x:.1f}"
    bal_rows.append({"o": name, "c": t(vals[0]), "v1": t(vals[1]), "v2": t(vals[2])})
bal_data = {"cols": [{"k": "o", "l": "About the parents"}, {"k": "c", "l": "Control", "num": True}, {"k": "v1", "l": "Video 1", "num": True}, {"k": "v2", "l": "Video 2", "num": True}], "rows": bal_rows} if N_C >= 3 else None

bar_arm = {"rows": [{"g": "Study group", "l": ARM_LAB[v], "m": r1(m_arm[v], 0), "n": n_arm[v], "slot": ARM_SLOT[v]} for v in (0, 1, 2) if m_arm[v] is not None], "unit": "Rs"}
bar_ord = {"rows": [], "unit": "Rs"}
for arm in (0, 1, 2):
    for o in (1, 2):
        ser = C.loc[arm_masks[arm] & ords[o], "_bid"].dropna()
        if len(ser) >= 3:
            bar_ord["rows"].append({"g": ARM_LAB[arm], "l": "Asked before own bid" if o == 1 else "Asked after own bid", "m": r1(float(ser.mean()), 0), "n": int(len(ser)), "slot": ARM_SLOT[arm]})
bar_arm = bar_arm if bar_arm["rows"] else None
bar_ord = bar_ord if bar_ord["rows"] else None

_best_arm = max((v for v in (0, 1, 2) if m_arm[v] is not None), key=lambda v: m_arm[v], default=None)
te_ins = ([{"cls": "teal", "html": f"<strong>Highest average bid:</strong> <strong>{ARM_LAB[_best_arm]}</strong> at <strong>{fmt_rs(m_arm[_best_arm])}</strong> "
            + " · ".join(f"{ARM_LAB[v]} {fmt_rs(m_arm[v])}" for v in (0, 1, 2) if m_arm[v] is not None and v != _best_arm) + ". Treat as indicative until fieldwork is complete."}] if _best_arm is not None else [])
PANELS.append({
    "id": "effects", "tab": "🧪 Video Effects", "eyebrow": "Section 07 · Randomised information experiment", "title": "Did the Videos Make a Difference?",
    "blurb": "Households were randomly assigned to no video (Control), Video 1 or Video 2 before the willingness-to-pay exercise, and to hearing about other parents' behaviour before or after their own bid. Compare the groups below.",
    "seg": False, "blocks": [
        {"t": "insight_static", "cls": "navy", "html": "<strong>🧪 Design:</strong> <strong>Control</strong> sees no video. <strong>Video 1</strong> and <strong>Video 2</strong> are two short films about classroom air purifiers shown before the bid. Assignment is pre-drawn in the sampling frame with a fixed seed, so group is unrelated to who the enumerator visits."},
        {"t": "kpis", "d": te_kpis},
        {"t": "grid", "cols": 2, "cards": [
            card("hbar", "teArmCi", "Average amount parents would pay, by study group", "Average final bid in Control, Video 1 and Video 2", bar_arm, opt={"grouped": True}),
            card("hbar", "teOrder", "Does asking about other parents first change the bid?", "Average bid when parents were asked about other parents before or after their own bid", bar_ord, opt={"grouped": True}, auto=True)]},
        {"t": "grid", "cols": 1, "cards": [
            card("table", "teEff", "Did the videos change what parents say?", "Average for each group. The highest number in each row is highlighted.", eff_data, full=True, opt={"noSort": True})]},
        {"t": "grid", "cols": 1, "cards": [
            card("table", "teBal", "Are the three groups alike?", "Background of the parents in each group. Similar numbers mean a fair comparison.", bal_data, full=True, opt={"noSort": True})]},
        {"t": "insights", "d": te_ins, "row": False},
        {"t": "note", "html": "Figures use completed interviews only, counted once per household. Group sizes are roughly a third of the sample each."},
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
        K("⏳", f"{med(disc):.0f}%" if med(disc) is not None else "—", "Extra Wanted to Wait a Year", "Typical yearly return parents ask for (median)", "neutral", "navy"),
        K("🐢", r1(avg(num(df, "tp_qual"))) or "—", "Patience Score (0–10)", "10 = very willing to give up something today", "up", "teal"),
        K("🛑", fmt_p(sh(num(df, "tp1"), lambda s: s == 1)), f"Take Rs {TP_BASE:,} Today over the larger amount", "First question, one year later", "neutral", "amber"),
        K("✔️", fmt_p(sh(num(df, "tp_check"), lambda s: s == 1)), "Consistent Check Answers", f"Prefer {TP_BASE:,} today to {TP_BASE:,} in a year", "up", "green"),
        K("📲", fmt_rs(avg(num(df, "topup_amount"))), "Mean Mobile Top-up", f"Rs {TOPUP_BASE:,} per child less any contribution paid", "navy", "purple"),
    ]


def time_ins(df):
    return [{"cls": "navy", "html": f"<strong>Impatience:</strong> the typical parent demands about <strong>{f'{med(df['_disc']):.0f}%' if med(df['_disc']) is not None else '—'}</strong> a year to give up money today, a very high implicit rate that is common in low-income settings and matters for how much weight future health benefits receive."},
            {"cls": "teal", "html": f"<strong>Data quality:</strong> <strong>{fmt_p(sh(num(df, 'tp_check'), lambda s: s == 1))}</strong> pass the consistency check ({TP_BASE:,} today vs the same amount in a year); the remainder may not have understood the question."}]


PANELS.append({
    "id": "time", "tab": "⏳ Time & Payout", "eyebrow": "Section 09 · Time preferences & incentive payout", "title": "Time Preferences & Mobile Top-up",
    "blurb": f"Three adaptive choices between Rs {TP_BASE:,} today and a larger amount in one year place each respondent's patience on a common scale. Participants also receive a mobile top-up worth Rs {TOPUP_BASE:,} per child less anything they contributed.",
    "seg": True, "blocks": [
        {"t": "kpis", "d": SD(time_kpis)},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "tpDisc", "Implied annual discount rate", "The reward, per year, at which each respondent switches from “today” to “in one year”", SD(f_hist(lambda d: d["_disc"], [-1000, 25, 50, 100, 200, 1e9], ["<25%", "25–49%", "50–99%", "100–199%", "200%+"])), var="tp_bin", opt={"color": 4, "fmt": "pct"}),
            card("bar", "tpQual", "Willingness to give up something today", "0 = completely unwilling, 10 = very willing", SD(lambda d: {"items": items([str(i) for i in range(11)], [int((num(d, "tp_qual") == i).sum()) for i in range(11)], int(num(d, "tp_qual").notna().sum())), "base": int(num(d, "tp_qual").notna().sum())} if num(d, "tp_qual").notna().any() else None), var="tp_qual", opt={"color": 6})]},
        {"t": "grid", "cols": 3, "cards": [
            card("donut", "tpFirst", f"First choice: Rs {TP_BASE:,} today or more in a year", "The opening question of the bisection", SD(f_dist("tp1", "timeopt", hide=(98,))), var="tp1", opt={"colors": [2, 3, "muted"]}),
            card("donut", "tpCheck", "Consistency check", f"Rs {TP_BASE:,} today or the same Rs {TP_BASE:,} in one year", SD(f_dist("tp_check", "timeopt", hide=(98,))), var="tp_check", opt={"colors": [6, 2, "muted"]}),
            card("donut", "tpNet", "Mobile network for top-up", "Network the respondent uses", SD(f_dist("mobile_network", "network_pk", drop_zero=True)), var="mobile_network", opt={"colors": [1, 2, 3, 4, "muted"]})]},
        {"t": "grid", "cols": 2, "cards": [
            card("bar", "tpTop", "Top-up amount", f"Rs {TOPUP_BASE:,} per child less the contribution paid under the random-price rule", SD(f_hist(lambda d: num(d, "topup_amount"), [0, 1000, 1500, 1800, 2000, 2001], ["<1,000", "1,000–1,499", "1,500–1,799", "1,800–1,999", "2,000"])), var="topup_amount", opt={"color": 5, "fmt": "pct"}),
            card("bar", "tpPaid", "Contribution actually paid", "Under the random-price rule; zero if the price drawn was above the bid", SD(f_hist(lambda d: d["_paid"], [-1, 1, 251, 501, 1001, 2001], ["Rs 0", "1–250", "251–500", "501–1,000", "1,001+"])), var="you_paid_amount_main", opt={"color": 3, "fmt": "pct"})]},
        {"t": "insights", "d": SD(time_ins), "row": True},
        {"t": "note", "html": "The implied discount rate is read from the final node of the three-step bisection (Rs 2,100 to Rs 8,000 offered in one year). Responses marked invalid or refused are excluded. Mobile numbers are collected for the top-up only and are never loaded into this dashboard."},
    ]})

# ====================================================================== PANEL 10 - HOUSEHOLD TRACKER
ST_LAB = {1: "Completed", 2: "Partial", 3: "Refused", 4: "Refused", 5: "Refused", 6: "Refused", 7: "Not reached", 8: "Not reached", 9: "Ineligible", 10: "Not reached"}
tr_rows = []
if FRAME_IS_POOL:                                         # only households that have actually been visited (the ID pool is not a visit list)
    ids = sorted(int(h) for h in D["_hh"].dropna().unique() if int(h) in frame.index)
else:
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
_hh_ck = D.dropna(subset=["_ckey"]).groupby("_hh")["_ckey"].last().to_dict()
for r in tr_rows:
    _ck = frame.loc[r["id"], "_ckey"] if HAS_FRAME_CLASS and r["id"] in frame.index else _hh_ck.get(r["id"])
    r["sc"] = _short(str(_ck).split("|")[0]) if isinstance(_ck, str) else "—"
    r["cl"] = CLS_LABEL.get(_ck, "—") if isinstance(_ck, str) else "—"
st_counts = {}
for r in tr_rows:
    st_counts[r["st"]] = st_counts.get(r["st"], 0) + 1
tr_kpis = [K("🏠", f"{len(tr_rows):,}", "Households Visited" if FRAME_IS_POOL else "Households in Frame",
             "Households with at least one attempt" if FRAME_IS_POOL else f"{ARM_LAB[0]} {arm_target[0]} · {ARM_LAB[1]} {arm_target[1]} · {ARM_LAB[2]} {arm_target[2]}", "navy", "navy"),
           K("✅", st_counts.get("Completed", 0), "Completed", f"{pct(st_counts.get('Completed', 0), len(tr_rows), 0):.0f}% of households visited" if FRAME_IS_POOL else f"{pct(st_counts.get('Completed', 0), len(tr_rows), 0):.0f}% of frame", "up", "green"),
           K("🔁", st_counts.get("Not reached", 0) + st_counts.get("Partial", 0) + n_ph_retry, "To Try Again",
             f"{n_ph_retry} parents whose phone was off or did not pick up · {st_counts.get('Not reached', 0) + st_counts.get('Partial', 0)} field revisits" if n_ph_nr else "Not reached or partial", "neutral", "amber"),
           K("🚫", st_counts.get("Refused", 0) + st_counts.get("Ineligible", 0) + n_ph_final, "Refused or Wrong Number",
             f"Refused, wrong or invalid number: no further attempt · {n_ph_final} from phone calls" if n_ph_nr else "No further visit planned", "down", ""),
           K("🕓", st_counts.get("Not visited", 0), "Not Yet Visited", "Waiting for a first attempt", "neutral", "teal")]
if FRAME_IS_POOL:
    tr_kpis = tr_kpis[:4]
PANELS.append({
    "id": "tracker", "tab": "🏠 Household Tracker", "eyebrow": "Section 10 · Operational tracker", "title": "Household Completion Tracker",
    "blurb": ("Status of every household visited so far: which are complete and which need a revisit. " if FRAME_IS_POOL else
              "Status of every household in the sampling frame: which are complete, which need a revisit and which have not been visited yet. ") +
             "Search by household ID, filter by status, arm or enumerator, click a column to sort, or export the current view to CSV.",
    "seg": False, "blocks": [
        {"t": "kpis", "d": tr_kpis},
        {"t": "blocks", "head": "📐 Progress by assignment block", "d": blocks, "foot": ("Each card is one study arm and question-order combination. Target = an equal share (one sixth) of the parents in the classes seen so far (it grows as new classes start); completed = interviews done. This is the quickest check that all six blocks fill evenly." if FRAME_IS_POOL else
                                                                                           "Each card is one study arm and question-order combination from the sampling frame. Target = households assigned to the block; completed = interviews done. This is the quickest check that all six blocks fill evenly.")},
        {"t": "tracker", "d": {"rows": tr_rows, "enums": sorted({r["e"] for r in tr_rows if r["e"] != "—"}), "arms": {str(k): v for k, v in ARM_LAB.items()}, "schools": cls_schools, "classes": [r["label"] for r in class_rows]}},
        {"t": "note", "html": "One row per household ID visited. Status reflects the latest attempt, or Completed if any attempt was completed. Only the household ID, assigned arm and visit outcome are shown: no names, addresses or phone numbers."},
    ]})

# ---- phone survey tab (kept apart from the field interviews)
if PH["calls"]:
    _pd = PH["disp"]
    _ph_kp = [K("📞", f"{PH['calls']:,}", "Phone Calls Made", f"{PH['parents']} unique parents called · {PH['calls'] - PH['parents']} repeat call(s)", "navy", "navy"),
              K("✅", f"{PH['agreed']:,}", "Agreed to Interview", f"{pct(PH['agreed'], PH['parents'], 0):.0f}% of parents called", "up", "green"),
              K("🔁", f"{PH['retry']:,}", "To Call Again", "Phone off, no answer or appointment", "neutral", "amber"),
              K("🚫", f"{PH['final']:,}", "Refused or Wrong Number", "No further call planned", "down", "")]
    _ph_dist = sorted(({"l": l, "n": n, "p": pct(n, max(PH["nr"], 1)), "lab": f"{n:,} parents"} for l, n in _pd.items()), key=lambda x: -x["n"])
    _ph_sch = [{"sc": k or "—", "nr": v["nr"], "retry": v["retry"], "final": v["final"]} for k, v in sorted(PH["school"].items())]
    _ph_day = [{"l": d.strftime("%b %d").replace(" 0", " "), "n": int(n)} for d, n in sorted(PH["by_day"].items())]
    _ph_en = [{"e": r["_en"] if "_en" in r else r.get("e", "—"), "calls": int(r["calls"]), "parents": int(r["parents"]), "ok": int(r["ok"])} for r in PH["by_enum"]]
    overview["blocks"][-1:-1] = [
            {"t": "note", "html": "<strong>📞 Phone Survey</strong> · parents are called first and the ones who agree are visited. Kept apart from the field interviews above: phone calls are never counted in the completed interviews, attempts or non-response."},
            {"t": "kpis", "d": _ph_kp},
            {"t": "grid", "cols": 2, "cards": [
                card("hbar", "phOut", "Phone Call Outcome", "Final outcome per parent (a parent called several times counts once; an agreed call wins, otherwise the latest call)", {"items": _ph_dist, "base": PH["nr"]}, opt={"color": 2, "fmt": "n"}),
                card("bar", "phDay", "Calls per Day", "Phone calls made on each day", {"items": [{"l": x["l"], "n": x["n"], "p": pct(x["n"], max(PH["calls"], 1))} for x in _ph_day], "base": PH["calls"]}, opt={"color": 1, "fmt": "n"})]},
    ]

# ---- tabs removed on request (Field Ops, Fairness, Patience)
PANELS = [p for p in PANELS if p["id"] not in ("ops", "fair", "time")]

# ====================================================================== DATA DOWNLOAD (raw CSV, in the Overview)
# The client downloads the field data exactly as delivered: with only a CSV in ..\Data the file is passed through
# byte for byte; with a .dta (or both) it is written out with the export's own column names, every column and row.
# It travels inside the AES-encrypted payload, so only someone with the dashboard password can open or download it.
# Set "raw_download_include_pii": false in the local config to strip identifier columns instead.
RAW = None
if RAW_DOWNLOAD:
    _last = RAW_DF["starttime"].replace("", np.nan).dropna().max() if "starttime" in RAW_DF else ""
    if csv_path and RAW_INCLUDE_PII and not (dta_path and PREFER == "dta"):
        # the SurveyCTO CSV itself, every column and row untouched (also when a .dta sits beside it, unless the .dta holds extra submissions)
        _csv = csv_path.read_text(encoding="utf-8-sig")
        _ncols = len(pd.read_csv(csv_path, dtype=str, nrows=0, encoding="utf-8-sig").columns)
    else:
        _src = RAW_DF
        if dta_path and PREFER == "dta":                # download = the .dta exactly (its rows, its columns, its corrected values)
            _src = RAW_DF[RAW_DF["_key"].isin(dta_df["_key"])]
            _keep = [c for c in dta_df.columns if not c.startswith("_") and (RAW_INCLUDE_PII or not is_pii(c))]
        else:
            _keep = [c for c in RAW_DF.columns if not c.startswith("_") and (RAW_INCLUDE_PII or not is_pii(c))]
        _raw_out = _src[_keep].rename(columns=lambda c: ORIG_NAMES.get(c, c))
        _csv = _raw_out.to_csv(index=False)
        _ncols = _raw_out.shape[1]
    RAW = {"name": f"AirQuality_field_data_{datetime.now().strftime('%Y-%m-%d')}.csv", "csv": _csv,
           "rows": int(len(RAW_DF)), "cols": int(_ncols), "last": str(_last)[:16] if _last else "-", "pii": RAW_INCLUDE_PII}
    say(f"Raw data  : {RAW['rows']} submission(s) x {RAW['cols']} columns, {len(_csv) / 1048576:.1f} MB CSV "
        f"({'all columns, nothing removed' if RAW_INCLUDE_PII else 'identifier columns removed'})")
    overview["blocks"].insert(2, {"t": "download"})
    # phone-call data, straight from its .dta (all calls, all columns), as a second download
    RAW_PH = None
    _pdta = next(iter(sorted(PHONE_DIR.glob("*Call Survey*.dta"))), None)
    if _pdta:
        _pd_ = load_file(_pdta)
        _pk = [c for c in _pd_.columns if not c.startswith("_") and (RAW_INCLUDE_PII or not is_pii(c))]
        _pcsv = _pd_[_pk].rename(columns=lambda c: ORIG_NAMES.get(c, c)).to_csv(index=False)
        _plast = _pd_["starttime"].replace("", np.nan).dropna().max() if "starttime" in _pd_ else ""
        RAW_PH = {"name": f"AirQuality_phone_calls_{datetime.now().strftime('%Y-%m-%d')}.csv", "csv": _pcsv,
                  "rows": int(len(_pd_)), "cols": len(_pk), "last": str(_plast)[:16] if _plast else "-"}
        say(f"Phone data: {RAW_PH['rows']} call(s) x {RAW_PH['cols']} columns from {_pdta.name}")
for _r in qa_rows:
    if _r["n"]:
        say(f"  check: {_r['name']}: {_r['n']} record(s)")

# ====================================================================== plain-language layer
_TEXT = TX.card_text(N_TARGET)
GUIDE = []


def _all_cards():
    for p in PANELS:
        for b in p["blocks"]:
            if b["t"] == "grid":
                for c in b["cards"]:
                    yield p, c


for _i, p in enumerate(PANELS, start=1):
    p["eyebrow"] = re.sub(r"^Section \d+", f"Section {_i:02d}", p["eyebrow"])
    p["tab"] = TX.TABS.get(p["id"], p["tab"])
    p["short"] = TX.SHORT.get(p["id"], "")
    if p["id"] in TX.PANEL_TITLE:
        p["title"] = TX.PANEL_TITLE[p["id"]]

_TPL = {"faPrio": "Ranked first or second most often: <strong>{l}</strong> ({p:.0f}% of parents)."}
for p, c in _all_cards():
    t = _TEXT.get(c["id"])
    if t:
        c["title"], c["desc"], _ur, hk = t
        if hk:
            c["help"] = hk
    else:
        _ur = ""
    if c["id"] in _TPL:
        c.setdefault("opt", {})["tpl"] = _TPL[c["id"]]
    dd = c["d"]
    targets = list(dd["_s"].values()) if isinstance(dd, dict) and "_s" in dd else [dd]
    for v in targets:
        if isinstance(v, dict) and "read" not in v:
            rd = TX.auto_read(c["kind"], v, c.get("opt"), c["id"])
            if rd:
                v["read"] = rd
    GUIDE.append({"panel": p["title"], "tab": p["tab"], "id": c["id"], "kind": c["kind"], "title": c["title"], "desc": c["desc"],
                  "var": c.get("var", ""), "ur": _ur, "help": c.get("help", "")})
(HERE / "build_cards.json").write_text(json.dumps({"panels": [{"id": p["id"], "title": p["title"], "eyebrow": p["eyebrow"], "blurb": p["blurb"]} for p in PANELS], "cards": GUIDE}, ensure_ascii=False, indent=1), encoding="utf-8")
DATA_HELP = {"help": TX.HELP, "generic": TX.GENERIC}

# ====================================================================== payload, PII scan, encrypt
DATA = {"meta": META, "segments": SEG_META, "panels": PANELS, "help": TX.HELP, "generic": TX.GENERIC}


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
payload = payload.replace("@PCT@", PCTM).replace("@CAP@", CAPS)          # tokens used in dashboard_text.py
# last-line PII scan of the serialised payload
_bad = re.findall(r"\b03\d{9}\b|\b\+?92\d{10}\b|[\w.]+@[\w.]+\.\w+", payload)
_bad = [b for b in _bad if "example" not in b]
if _bad:
    sys.exit(f"PII pattern found in payload ({_bad[:3]}...): build stopped.")
if RAW and RAW_PH:
    RAW["phone"] = RAW_PH
if RAW:                                                  # the downloadable file is added after the aggregate scan (it is the export itself)
    DATA["raw"] = RAW
    payload = json.dumps(clean(DATA), ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    payload = payload.replace("@PCT@", PCTM).replace("@CAP@", CAPS)
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
