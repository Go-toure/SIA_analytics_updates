# SIA analytics updates

Standalone R script that produces the regional SIA performance review from the
AFRO SIA Dashboard data (administrative coverage + LQAS): Excel workbook, PNG charts,
PowerPoint deck and an HTML/Markdown narrative.

No data is stored in this repository. Outputs are written to the output folder only.

## Run

```
Rscript sia_executive_review.R --from 2026-04-01 --to 2026-09-30 --zoom "Togo,Mali,Nigeria,Angola,Guinea,South Sudan"
```

Options: `--admin`, `--lqas` (input files), `--out` (output folder), `--from`, `--to`, `--zoom`, `--title`.
Without `--from/--to` the previous and current quarters of the latest data are used.

Required packages: dplyr, tidyr, stringr, lubridate, ggplot2, scales, purrr, tibble, forcats,
readxl, openxlsx, officer, flextable, ragg.

Standards (thresholds) are working defaults in the `CFG` block at the top of the script
and must be confirmed by the technical team.

The deck, workbook and charts also include a view by campaign round: one row per round of each country's latest assessable campaign, the four standards, and the children vaccinated taken from the administrative data (column TotalNbrVaccPolio, summed over the districts of the round). Files: `charts/04b_indicator_heatmap_by_round.png` (and `04c_by_round_part1/2.png` for the slides) and the sheet `Indicators_by_round`.

Report dates that fall before the campaign (shown as "date error") are corrected when the cause is a wrong year: the campaign calendar (lookup.xlsx) must confirm the campaign date, and the report date is then moved to the next same day and month on or after the campaign. Every change is listed in the sheet `Report_date_fixes`; report dates the calendar cannot confirm are left as they are.

## Deep check of the admin and SCOPE data (`py/admin_scope_deep_check.py`)

```
python py/admin_scope_deep_check.py
```

Reads `scope/data/AFRO_admin_data.xlsx`, `scope/data/scope_.csv` and the LQAS repository, places district names on AFROGDB through the crosswalk, and writes summaries to `deep_check/` (the row-level file `admin_record_flags.csv` stays local). It checks: the stored CVPolio against the official coverage (vaccinated / PopPolio) and what that does to the dashboard bands, extreme coverage, jumps of the target or vaccinated count between campaigns of the same district, conflicting duplicate rows, impossible dates, rounded counts (share ending in 00), a model that marks the 1% most unusual rows, whether every SCOPE planned district-round has an admin row, and how admin coverage agrees with LQAS in the same district and campaign date. Ethiopia is not reliable here (only 27% of its admin names can be placed on AFROGDB).
