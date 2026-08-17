# UCITS Compliance Rules (Generic Reference)

This skill performs an automated UCITS compliance pre-check on the portfolio
holdings supplied in `holdings.csv`. All thresholds are configurable in
`config.json` → `ucits`. Replace the example values with your fund's mandate.

## Checks performed

| # | Rule | Config key | Example default |
|---|------|-----------|-----------------|
| 1 | Stock count within range | `stock_count_min` / `stock_count_max` | 23–27 |
| 2 | Eligible-region only (e.g. Greater China) | `greater_china_only` | true |
| 3 | Single-name concentration ≤ cap | `single_name_cap_pct` | 10% |
| 4 | 5/40 rule: sum of names >5% ≤ 40% | `rule_5_40` | true |
| 5 | A-share exposure < cap | `a_share_cap_pct` | 30% |
| 6 | Taiwan exposure < cap | `taiwan_cap_pct` | 40% |
| 7 | Cash < cap | `cash_cap_pct` | 20% |
| 8 | Flagged names (domicile violations) | `flagged_names` | [] |

## How the checker reads `holdings.csv`

Required columns:

```
name,region,weight,latest,entry,target,compliance
```

- `region` must be one of: `HK`, `A-share`, `Taiwan`, `Cash`, `Other`.
- `weight` is the position weight in percent (e.g. `8.0`).
- Rows with `region = Cash` are excluded from the stock count and treated as cash.
- Any `region` not in `{HK, A-share, Taiwan, Cash}` is treated as **non-eligible**
  and triggers the "eligible-region only" violation.
- Names listed in `ucits.flagged_names` are explicitly flagged with the
  `flagged_note` message (e.g. a name domiciled outside the eligible region).

## Region caps

The checker sums weights by `region`:
- `A-share` sum is compared to `a_share_cap_pct`.
- `Taiwan` sum is compared to `taiwan_cap_pct`.
- `Cash` sum is compared to `cash_cap_pct`.

## Notes

- This is a **pre-check /sanity check**, not a substitute for official compliance
  sign-off.
- Keep `region` labels consistent (case-insensitive match). Unknown labels fall
  into the "non-eligible" bucket and will flag.
