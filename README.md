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
