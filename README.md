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
