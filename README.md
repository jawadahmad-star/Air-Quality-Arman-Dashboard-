# Air Quality Project: Survey Dashboard

Client dashboard for the Research Solutions air quality household survey
(parents' perceptions of air pollution and willingness to pay for classroom air
purifiers, Lahore).

**Live:** https://airqualityproject.rs.org.pk  ·  access is by password (held by Research Solutions).

## What is in this repo

| File | Role |
|---|---|
| `index.html` | The built dashboard. The only file GitHub Pages serves. All data inside is **AES-256-GCM encrypted**; it is unreadable without the password. |
| `dashboard_template.html` | Layout, charts and password gate. |
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

## Hosting

GitHub Pages, branch `main`, folder `/ (root)`, custom domain `airqualityproject.rs.org.pk`, Enforce HTTPS.
DNS: a `CNAME` record for `airqualityproject` pointing to `jawadahmad-star.github.io`.

Research Solutions (M&A Research Solutions LLC) · www.rs.org.pk
