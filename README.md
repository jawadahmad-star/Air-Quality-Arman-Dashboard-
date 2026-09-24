# Air Quality Project: Survey Dashboard

Client dashboard for the Research Solutions air quality household survey
(parents' perceptions of air pollution and willingness to pay for classroom air
purifiers, Lahore).

**Live:** https://airqualityproject.rs.org.pk  ·  access is by password (held by Research Solutions).

## What is in this repo

| File | Role |
|---|---|
| `index.html` | The built dashboard. The only file GitHub Pages serves. All data inside is **AES-256-GCM encrypted**; it is unreadable without the password. |
| `dashboard_template.html` | Layout, charts, navigation and password gate. |
| `dashboard_text.py` | Plain-language chart titles, descriptions and "how to read this" notes. |
| `build_dashboard.py` | Reads the SurveyCTO export, computes every aggregate, encrypts and writes `index.html`. |
| `update_dashboard.bat` | One-click refresh: rebuild, check nothing sensitive is staged, commit, push. |
| `CNAME`, `.nojekyll` | GitHub Pages custom domain and config. |
| `dashboard.config.example.json` | Template for the local, git-ignored `dashboard.config.json` (holds the password). |

Raw data, identified exports, the frame and the password are **never** committed (`.gitignore`).

## Daily update

1. Export from SurveyCTO into `..\Data\`: the wide **CSV** and/or the Stata **.dta** (either, or both).
2. Double-click `update_dashboard.bat`.

The builder then:

- loads the newest `.csv` and `.dta` in `..\Data\`. If both are present they are reconciled row by row on the
  SurveyCTO `KEY`; where a value differs, the cleaned `.dta` wins (`"prefer"` in the config). Counts of matched,
  one-sided and conflicting rows are shown in **Field Operations → Data Source Check**;
- drops names, phone numbers, addresses, device ids and GPS points on load;
- counts a household once (latest completed submission), and only if its ID is in the assignment frame;
- switches the dashboard from **DEMO** to **LIVE** automatically once real data is present.

With nothing in `..\Data\`, the builder uses the synthetic files in `..\Data\_dummy\` and stamps the page **Demo**.

`python build_dashboard.py --check` validates the data and writes nothing.

## Setup on a new machine

```
pip install pandas numpy openpyxl cryptography pyreadstat
copy dashboard.config.example.json dashboard.config.json   (then set the password)
```

## Study design the dashboard follows

Government schools → classes 4, 5 and 6 → the parents of each class (roughly 1,200–1,300 households; the number of schools and classes can be anything, the dashboard adapts).
Parents are phoned from the school list and visited by appointment. A classroom gets a purifier only if at least **30% of that class's own parents**
end up contributing (random price at or below their bid; parents not interviewed do not count). The **Classrooms** tab applies that rule class by class.

### Sampling frame (prefill) columns

`prefill_data.xlsx` should carry, besides `hh_id`, `order_key`, `ap_lm_arm`: **`school`**, **`grade`** (4/5/6), optional **`section`** (A/B) and **`class_size`**.
The dashboard uses them to build classes and targets. If they are missing it falls back to the `school_child` and `grade_child` text typed by enumerators,
which is fragile (spelling differences split one class in two), so adding them to the frame is strongly recommended.

## Hosting

GitHub Pages, branch `main`, folder `/ (root)`, custom domain `airqualityproject.rs.org.pk`, Enforce HTTPS.
DNS: a `CNAME` record for `airqualityproject` pointing to `jawadahmad-star.github.io`.

Research Solutions (M&A Research Solutions LLC) · www.rs.org.pk
