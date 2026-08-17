---
name: ucits-outlook-report
description: Generate a UCITS / Greater-China Outlook email deep-analysis report (.docx) from fetched broker emails. Fetches research emails (no link crawling), then builds a 5-part Word report (Major Events / Investment Summary / Long Ideas / Portfolio Table + automated UCITS compliance check / Link Status). Fully parameterized — no hardcoded fund, holdings, or login data. Use when asked to "make the daily UCITS report", "turn Outlook research emails into a portfolio report", "check UCITS compliance on holdings", or "build the broker-email digest".
---

# UCITS Outlook Email Report Skill

Turn a day's worth of broker research emails (Daiwa / DBS / UOBKH / Mizuho / etc.)
into a structured, compliance-checked Word report.

## What it does

1. **Fetch** research emails from Outlook (Playwright, persistent Edge profile, no link
   tracking — avoids browser hang on close).
2. **Generate** a 5-part `.docx`:
   - Cover (fund name, date, source stats)
   - Part 1 — Major Market Events (four-part blocks: Time/Source · Key Data · Analysis)
   - Part 2 — Investment Ideas Summary
   - Part 3 — Long Ideas
   - Part 4 — Portfolio Allocation table **+ automated UCITS compliance pre-check**
   - Part 5 — Link Status

## Why it is reusable (no secrets, no real data)

- **Holdings are blanked**: the generator reads `holdings.csv` — ship the example
  (`examples/holdings_example.csv`) and let each user drop in their own (gitignored).
- **Login is blanked**: the fetcher takes the profile path via `--profile` / env var;
  the persistent profile (your saved Outlook session) is never committed (gitignored).
- **Fund name / thresholds / fonts / language** are all in `config.example.json`.

## Workflow

### 0. One-time setup
```bash
pip install python-docx playwright
playwright install msedge        # or use your installed Edge channel
```

### 1. Fetch emails (optional — you can also supply your own JSON)
```bash
# First run only: log into Outlook with Edge once so a persistent profile is created.
python scripts/fetch_emails.py --date 2026-08-13 \
    --profile /path/to/.edge_profile \
    --url https://outlook.office.com/mail/inbox
# -> writes reports/emails_2026-08-13.json
```

### 2. Configure
Copy `config.example.json` → `config.json` and fill in:
- `report.fund_name`, `report.manager`, `report.report_date`
- `ucits.*` thresholds for your mandate
- `paths.*` (or keep defaults — resolved relative to the config file)

### 3. Synthesize daily content (the AGENT does this step)
After fetching (step 1), read `reports/emails_<date>.json` — the raw emails you just pulled.
As the agent, you must SYNTHESIZE `events.json` yourself, following the schema in
`examples/events.example.json`:
- `cover_stats` — source count / total fetched / deduped / GC-relevant / non-GC numbers
- `part1_sections` — major market events; each has `importance`, `title`, `time_source`,
  `key_data`, `analysis`
- `part2_summary` — investment ideas summary
- `part3_long_ideas` — long ideas
- `part4_note` — portfolio-table note
- `part5_link_status` — link-crawl status (usually "not crawled")
The skill ships NO opinions — only the structure. If you are handed only the raw emails
and told "make the report", producing this `events.json` from them IS your job.
Also prepare `holdings.csv` — current positions (columns:
`name,region,weight,latest,entry,target,compliance`).

### 4. Generate
```bash
python scripts/generate_report.py config.json
# -> examples/sample_report.docx (or your configured output path)
```

## Input file reference

| File | Purpose | Committed? |
|------|---------|-----------|
| `config.example.json` | Template config (all params) | ✅ yes |
| `examples/holdings_example.csv` | Example portfolio (fictional) | ✅ yes |
| `examples/events.example.json` | Example daily analysis | ✅ yes |
| `config.json` | Your real config | ❌ gitignored |
| `holdings.csv` | Your real positions | ❌ gitignored |
| `.edge_profile/` | Your saved Outlook login | ❌ gitignored |

## Compliance check

The generator auto-computes, from `holdings.csv`:
stock count, eligible-region-only, single-name cap, 5/40 rule, A-share / Taiwan /
cash caps, and flagged-name detection. See `references/ucits_compliance.md`.

> This is a **pre-check**, not a compliance sign-off.

## Notes / gotchas

- Outlook's DOM selectors can change; if the fetcher returns empty bodies, re-inspect
  `body_sels` in `scripts/fetch_emails.py`.
- The fetcher is headless=False (needs a visible browser for the saved session). Run it
  on a machine where you can log in once.
- All analysis text lives in `events.json` — the skill ships no opinions, only structure.
  The agent synthesizes `events.json` from the fetched raw emails (`reports/emails_*.json`);
  the skill does not generate analysis content on its own.
