# UCITS Outlook Email Report (reusable skill)

A self-contained, **zero-secret** skill that turns a day of broker research emails
into a structured, compliance-checked Word report. Everything fund-specific
(holdings, fund name, login session) is parameterized or gitignored — so you can
fork it, push it to a **public** GitHub repo, and let others reuse it safely.

## What's inside

```
ucits-outlook-report/
├── SKILL.md                  # skill metadata + workflow
├── README.md                 # this file
├── .gitignore                # excludes real config / holdings / login profile
├── config.example.json       # all parameters (fund, thresholds, fonts, paths)
├── scripts/
│   ├── fetch_emails.py       # Outlook fetcher (no link crawling)
│   └── generate_report.py    # parameterized .docx generator
├── examples/
│   ├── holdings_example.csv  # fictional portfolio (shows CSV format)
│   ├── events.example.json   # example daily analysis content
│   └── sample_report.docx    # generated sample (safe to commit)
└── references/
    └── ucits_compliance.md   # compliance rule reference
```

## Quick start

```bash
pip install python-docx playwright
playwright install msedge

# 1) Fetch (first run logs in once to create the persistent profile)
python scripts/fetch_emails.py --date 2026-08-13 --profile /path/to/.edge_profile

# 2) Copy & edit config
cp config.example.json config.json
#   edit fund_name, manager, report_date, ucits.* thresholds

# 3) Drop in your daily content
#   - events.json  (from examples/events.example.json)
#   - holdings.csv (columns: name,region,weight,latest,entry,target,compliance)

# 4) Generate
python scripts/generate_report.py config.json
```

## Pushing to GitHub

This repo is safe to publish:

- ✅ No real fund name, no real holdings, no credentials.
- ✅ `.gitignore` blocks `config.json`, `holdings.csv`, `events.json`, `.edge_profile/`.
- ✅ Only the `*example*` files and `sample_report.docx` (fictional data) are committed.

```bash
git init
git add .
git commit -m "Add reusable UCITS Outlook email report skill"
gh repo create ucits-outlook-report --public --source=. --push
```

Each user clones it, adds their own `config.json` + `holdings.csv` + `.edge_profile`,
and runs the two scripts. Done.

## Requirements

- Python 3.10+
- `python-docx`, `playwright` (+ `msedge` browser)
- A one-time manual Outlook login to create the persistent Edge profile
