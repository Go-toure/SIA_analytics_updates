#!/usr/bin/env Rscript
# =============================================================================
# WHO AFRO | SIA EXECUTIVE REVIEW  -  standalone analysis script
# -----------------------------------------------------------------------------
# What it does
#   Reads the AFRO administrative data (and, when present, the LQAS data) of the
#   AFRO SIA Dashboard and produces a regional review of SIA performance:
#     1. a score for every campaign round  (indicators met / indicators evaluable)
#     2. the latest score of every country, its trend (first vs latest campaign),
#        its performance band and its action segment
#     3. what drives the score (indicators and pillars, previous vs latest quarter)
#     4. data quality as an operational risk (not evaluable campaigns, anomalies)
#     5. a decision table (action, responsible, timeline, evidence)
#   and writes, in one folder:
#     SIA_review_results.xlsx      all tables
#     SIA_review_slides.pptx       12-slide executive review
#     SIA_review_narrative.html    key messages + tables, ready for the dashboard
#     SIA_review_narrative.md      same text in markdown
#     charts/*.png                 the charts
#     SIA_review_results.rds       the results object (for the Shiny app)
#
# How to run (Windows, R 4.1 or newer)
#   Rscript sia_executive_review.R
#   Rscript sia_executive_review.R --from 2026-04-01 --to 2026-09-30
#   Rscript sia_executive_review.R --admin "C:/path/AFRO_admin_data.xlsx" --out "C:/path/folder"
#   or from R / the Shiny app:   source("sia_executive_review.R"); res <- run_sia_review()
#
# Options (all optional)
#   --from / --to     period of the review (default: the last 6 months of the data)
#   --admin / --lqas  data files (.rds or .xlsx / .csv); default: the dashboard's input folder
#   --out             output folder (default: SIA_feedbacks)
#   --title           title shown on the slides
#
# Packages: dplyr tidyr stringr lubridate ggplot2 scales openxlsx officer flextable
#           purrr tibble forcats zip   (readxl only when the admin file is .xlsx)
#
# IMPORTANT - the standards below (CFG) are working defaults. They are written
# at the top of the script and in the README sheet of the workbook so that the
# technical team can confirm or change them in one place. Nothing else in the
# script needs to be edited.
# =============================================================================

suppressPackageStartupMessages({
  library(dplyr); library(tidyr); library(stringr); library(lubridate)
  library(ggplot2); library(scales); library(purrr); library(tibble); library(forcats)
})

`%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || (length(a) == 1 && is.na(a))) b else a

# ----------------------------------------------------------------------------- 1. SETTINGS
CFG <- list(
  out_dir        = "C:/Users/TOURE/Documents/SIA_feedbacks",
  dashboard_dir  = "C:/Users/TOURE/Documents/AFRO-SIA_Dashboard",
  admin_file     = NULL,            # NULL = look in <dashboard_dir>/input
  lqas_file      = NULL,
  period_from    = NULL,            # NULL = last `default_months` months of the data
  period_to      = NULL,
  default_months = 6,
  title          = "Regional Review of SIA Performance",
  zoom_countries = c("Togo", "Mali", "Nigeria", "Angola", "Guinea"),   # countries of the LQAS zoom chart
  org            = "WHO AFRO",

  # ---- standards (an indicator is met when its value reaches the standard) ----
  cov_std        = 95,    # % coverage a district / a campaign has to reach
  cov_ceiling    = 120,   # % above which a coverage figure is treated as implausible
  cov_low        = 80,    # % under which a district is counted as low coverage
  share_ge95_std = 90,    # % of districts that must reach the coverage standard
  share_low_max  = 5,     # max % of districts below `cov_low`
  complete_std   = 98,    # % of districts with both target and vaccinated figures
  implaus_max    = 10,    # max % of districts with an implausible figure (working default; coverage > ceiling is frequent, adjust to the technical standard)
  delay_max_days = 14,    # median days between the campaign and its report
  report_min_share = 50,  # % of districts with a report date needed to judge timeliness
  lqas_pass_std  = 90,    # % of LQAS lots that must pass
  informed_std   = 90,    # % of caregivers informed before the campaign
  noncomp_max    = 2,     # max % of sampled children missed because of refusal
  lqas_grace_days = 30,   # LQAS results are expected within this many days
  min_evaluable  = 4,     # fewer evaluable indicators = campaign not evaluable (N/E)
  full_evidence  = 9,     # evaluable indicators for "full" evidence (fewer = partial)

  # ---- bands, trends, segments ----
  bands          = c(40, 55, 75),   # below 40 | 40-54 | 55-74 | 75 or above
  trend_tol      = 2,               # points of change counted as "no change"
  weak_below     = 45,              # "persistently weak" = below this in both quarters
  change_flag    = 5,               # points of change that count as improving / worsening
  min_districts  = 3,               # smaller campaign rounds are not scored (mop-ups)

  # ---- decisions ----
  days_validate  = "10 working days",
  days_plan      = "15 working days and before the next SIA round"
)

COL <- list(navy = "#1B3A5F", green = "#2E7D32", gold = "#D9A441", orange = "#C77700",
            red = "#B00020", grey = "#B8C4D0", lightgrey = "#EEF2F6", text = "#1F2937",
            muted = "#6B7280", blue = "#0F6CBD", teal = "#14808A")

# ----------------------------------------------------------------------------- 2. INDICATORS
IND <- tribble(
  ~id,    ~pillar,              ~label,                           ~standard_text,
  "COV1", "Coverage",           "Campaign coverage \u2265 95%",    "Campaign coverage between the standard and the ceiling",
  "COV2", "Coverage",           "Districts at \u2265 95%",         "At least {share_ge95_std}% of districts reach the standard",
  "COV3", "Coverage",           "Few districts below 80%",         "At most {share_low_max}% of districts below {cov_low}%",
  "DQ1",  "Data quality",       "Complete district data",          "At least {complete_std}% of districts with target and vaccinated",
  "DQ2",  "Data quality",       "Plausible coverage figures",      "At most {implaus_max}% of districts above {cov_ceiling}%",
  "DQ3",  "Data quality",       "Timely reporting",                "Median report delay of at most {delay_max_days} days",
  "DQ4",  "Data quality",       "Consistent campaign code",        "Campaign code matches the country",
  "LQ1",  "LQAS verification",  "LQAS conducted",                  "LQAS results exist for the round",
  "LQ2",  "LQAS verification",  "LQAS accepted",                   "At least {lqas_pass_std}% of lots pass",
  "LQ3",  "LQAS verification",  "Caregivers informed",             "At least {informed_std}% of caregivers informed",
  "LQ4",  "LQAS verification",  "Refusals under control",          "At most {noncomp_max}% of sampled children missed through refusal"
)
IND$standard_text <- map_chr(IND$standard_text, function(s) {
  str_replace_all(s, "\\{(\\w+)\\}", function(m) as.character(CFG[[str_remove_all(m, "[{}]")]]))
})

# ----------------------------------------------------------------------------- 3. HELPERS
nice_country <- function(x) {
  x <- as.character(x)
  out <- str_to_title(x)
  out[str_detect(x, "DEMOCRATIC REPUBLIC")] <- "DRC"
  out[x == "CENTRAL AFRICAN REPUBLIC"] <- "CAR"
  out[str_detect(x, "IVOIRE")] <- "Côte d'Ivoire"
  out[str_detect(x, "^GUINEA.?BISSAU")] <- "Guinea-Bissau"
  out[x == "CONGO"] <- "Congo"
  out[x == "UNITED REPUBLIC OF TANZANIA"] <- "Tanzania"
  out
}
qlabel   <- function(d) paste0("Q", quarter(d), " ", year(d))
qkey     <- function(d) year(d) * 10 + quarter(d)
pct      <- function(x, d = 0) ifelse(is.na(x), "n/a", paste0(formatC(round(x, d), format = "f", digits = d), "%"))
round_num <- function(x) suppressWarnings(as.integer(str_extract(as.character(x), "\\d+")))
norm_resp <- function(x) toupper(str_replace_all(as.character(x), "[^A-Za-z0-9]", ""))
msg <- function(...) cat(format(Sys.time(), "%H:%M:%S"), "|", ..., "\n")

parse_args <- function(args) {
  out <- list(); i <- 1
  while (i <= length(args)) {
    if (str_starts(args[i], "--")) {
      key <- str_remove(args[i], "^--")
      if (i < length(args) && !str_starts(args[i + 1], "--")) { out[[key]] <- args[i + 1]; i <- i + 2 }
      else { out[[key]] <- TRUE; i <- i + 1 }
    } else i <- i + 1
  }
  out
}

read_any <- function(path) {
  ext <- tolower(tools::file_ext(path))
  if (ext == "rds") return(readRDS(path))
  if (ext %in% c("xlsx", "xlsm", "xls")) {
    if (!requireNamespace("readxl", quietly = TRUE)) stop("Install the 'readxl' package to read ", path)
    return(readxl::read_excel(path, guess_max = 50000))
  }
  if (ext == "csv") return(utils::read.csv(path, stringsAsFactors = FALSE, check.names = FALSE))
  stop("Unsupported file type: ", path)
}

# the newest existing file among the candidate folders / names
find_input <- function(explicit, names, dirs) {
  if (!is.null(explicit) && nzchar(explicit)) { if (file.exists(explicit)) return(explicit) else stop("File not found: ", explicit) }
  cand <- unlist(lapply(dirs, function(d) file.path(d, names)))
  cand <- cand[file.exists(cand)]
  if (!length(cand)) return(NA_character_)
  cand[which.max(file.mtime(cand))]
}

as_date_safe <- function(x) {
  if (inherits(x, "Date")) return(x)
  if (inherits(x, "POSIXt")) return(as.Date(x))
  suppressWarnings(as.Date(x))
}

# ----------------------------------------------------------------------------- 4. DATA
load_admin <- function(path) {
  a <- as_tibble(read_any(path))
  need <- c("Country", "Province", "District", "Response", "Round_Add", "SIA_date", "PopPolio", "TotalNbrVaccPolio")
  miss <- setdiff(need, names(a))
  if (length(miss)) stop("The admin file lacks the column(s): ", paste(miss, collapse = ", "))
  for (c in c("Admin_3", "Vaccine_type", "Date_Report", "AFRO_block")) if (!c %in% names(a)) a[[c]] <- NA
  a %>% transmute(
    Country = as.character(Country), Province = as.character(Province), District = as.character(District),
    Admin_3 = as.character(Admin_3), Response = str_squish(as.character(Response)),
    Round_Add = str_squish(as.character(Round_Add)), Vaccine_type = as.character(Vaccine_type),
    SIA_date = as_date_safe(SIA_date), Date_Report = as_date_safe(Date_Report),
    PopPolio = suppressWarnings(as.numeric(PopPolio)), Vacc = suppressWarnings(as.numeric(TotalNbrVaccPolio)),
    AFRO_block = as.character(AFRO_block))
}

load_lqas <- function(path) {
  if (is.na(path)) return(NULL)
  x <- as_tibble(read_any(path))
  g <- function(...) { n <- intersect(c(...), names(x)); if (length(n)) x[[n[1]]] else rep(NA, nrow(x)) }
  num <- function(v) suppressWarnings(as.numeric(v))
  out <- tibble(
    lq_country = as.character(g("country")), province = as.character(g("province")), district = as.character(g("district")),
    Response = str_squish(as.character(g("response", "Response"))), Round_Add = str_squish(as.character(g("roundNumber"))),
    lqas_date = as_date_safe(g("lqas_end_date", "end_date", "lqas_start_date")),
    sampled = num(g("total_sampled")), vaccinated = num(g("total_vaccinated")),
    status = str_to_lower(as.character(g("status"))),
    noncomp = coalesce(num(g("r_non_compliance")), num(g("r_Non_Compliance"))),
    informed = coalesce(num(g("prct_care_giver_informed_SIA")), num(g("percent_care_Giver_Informed_SIA"))))
  out <- out %>% filter(!is.na(Response), !is.na(status)) %>%
    distinct(lq_country, province, district, Response, Round_Add, lqas_date, sampled, vaccinated, .keep_all = TRUE) %>%
    mutate(pass = status %in% c("pass", "passed"), key = paste(norm_resp(Response), round_num(Round_Add), sep = "|"))
  out
}

# ----------------------------------------------------------------------------- 5. ANALYSIS
sum_na <- function(x) if (all(is.na(x))) NA_real_ else sum(x, na.rm = TRUE)
max_na <- function(x) if (all(is.na(x))) NA_real_ else max(x, na.rm = TRUE)
min_date <- function(x) { x <- x[!is.na(x)]; if (length(x)) min(x) else as.Date(NA) }
max_date <- function(x) { x <- x[!is.na(x)]; if (length(x)) max(x) else as.Date(NA) }

band_of <- function(s, b = CFG$bands) {
  lv <- c(sprintf("Below %d%%", b[1]), sprintf("%d–%d%%", b[1], b[2] - 1), sprintf("%d–%d%%", b[2], b[3] - 1), sprintf("%d%% or above", b[3]))
  factor(ifelse(is.na(s), NA, ifelse(s < b[1], lv[1], ifelse(s < b[2], lv[2], ifelse(s < b[3], lv[3], lv[4])))), levels = lv)
}
risk_of <- function(band) {
  r <- c("Very high", "High", "Moderate", "Low")[as.integer(band)]
  factor(r, levels = c("Very high", "High", "Moderate", "Low"))
}
band_colour <- function(band) c(COL$red, COL$orange, COL$gold, COL$green)[as.integer(band)]

# the usual country of every campaign-code prefix (ANG, MLW, DRC ...) over the whole file
prefix_table <- function(adm) {
  adm %>% filter(!is.na(Response)) %>%
    mutate(prefix = str_to_upper(str_extract(Response, "^[A-Za-z]+"))) %>%
    filter(!is.na(prefix)) %>% count(prefix, Country, name = "n") %>%
    group_by(prefix) %>% mutate(total = sum(n)) %>% slice_max(n, n = 1, with_ties = FALSE) %>% ungroup() %>%
    transmute(prefix, usual_country = Country, usual_share = n / total, prefix_rows = total)
}

build_units <- function(adm, from, to) {
  w <- adm %>% filter(!is.na(SIA_date), SIA_date >= from, SIA_date <= to, !is.na(Response), Response != "") %>%
    mutate(Round_Add = ifelse(is.na(Round_Add) | Round_Add == "", "Rnd?", Round_Add))
  u <- w %>% group_by(Country, Response, Round_Add, Province, District, Admin_3) %>%
    summarise(pop = max_na(PopPolio), vac = max_na(Vacc), sia_date = min_date(SIA_date), report = max_date(Date_Report),
              vaccine = paste(sort(unique(na.omit(Vaccine_type))), collapse = " + "), .groups = "drop") %>%
    group_by(Country, Response, Round_Add, Province, District) %>%
    summarise(pop = sum_na(pop), vac = sum_na(vac), sia_date = min_date(sia_date), report = max_date(report),
              vaccine = paste(sort(unique(vaccine[vaccine != ""])), collapse = " + "), .groups = "drop") %>%
    mutate(
      missing    = is.na(pop) | pop <= 0 | is.na(vac),
      cv         = ifelse(!missing, 100 * vac / pop, NA_real_),
      implausible = (!is.na(cv) & cv > CFG$cov_ceiling) | (!is.na(pop) & pop <= 0 & !is.na(vac) & vac > 0),
      low        = !is.na(cv) & cv < CFG$cov_low,
      ok95       = !is.na(cv) & cv >= CFG$cov_std & cv <= CFG$cov_ceiling,
      delay      = as.numeric(report - sia_date),
      key        = paste(Country, Response, Round_Add, sep = "|"))
  u
}

score_campaigns <- function(units, lqas, prefixes, lqas_max) {
  camp <- units %>% group_by(key, Country, Response, Round_Add) %>%
    summarise(
      n_districts = n(), date = median(sia_date, na.rm = TRUE), first_date = min_date(sia_date), vaccine = paste(sort(unique(vaccine[vaccine != ""])), collapse = " + "),
      pop_total = sum(pop[!missing], na.rm = TRUE), vac_total = sum(vac[!missing], na.rm = TRUE),
      n_missing = sum(missing), n_implausible = sum(implausible), n_low = sum(low), n_ok95 = sum(ok95),
      n_valid = sum(!missing), report_share = 100 * mean(!is.na(report)), med_delay = suppressWarnings(median(delay, na.rm = TRUE)),
      n_report_before = sum(!is.na(delay) & delay < 0), .groups = "drop") %>%
    mutate(med_delay = ifelse(is.finite(med_delay), med_delay, NA_real_),
           cov_total = ifelse(pop_total > 0 & n_valid > 0, 100 * vac_total / pop_total, NA_real_),
           share_ge95 = ifelse(n_valid > 0, 100 * n_ok95 / n_valid, NA_real_),
           share_low = ifelse(n_valid > 0, 100 * n_low / n_valid, NA_real_),
           complete = 100 * (n_districts - n_missing) / n_districts,
           implaus = 100 * n_implausible / n_districts,
           prefix = str_to_upper(str_extract(Response, "^[A-Za-z]+"))) %>%
    left_join(prefixes, by = "prefix") %>%
    mutate(code_ok = ifelse(is.na(usual_country) | prefix_rows < 20, NA, usual_country == Country),
           small = n_districts < CFG$min_districts)

  lq <- NULL
  if (!is.null(lqas) && nrow(lqas)) {
    lq <- lqas %>% group_by(key) %>% summarise(
      lq_lots = n(), lq_pass = 100 * mean(pass), lq_informed = suppressWarnings(mean(informed, na.rm = TRUE)),
      lq_noncomp = ifelse(sum(sampled, na.rm = TRUE) > 0 & any(!is.na(noncomp)), 100 * sum(noncomp, na.rm = TRUE) / sum(sampled[!is.na(noncomp)], na.rm = TRUE), NA_real_),
      lq_last = max_date(lqas_date), .groups = "drop") %>%
      mutate(lq_informed = ifelse(is.finite(lq_informed), lq_informed, NA_real_))
  }
  camp <- camp %>% mutate(lkey = paste(norm_resp(Response), round_num(Round_Add), sep = "|"))
  camp <- if (!is.null(lq)) left_join(camp, lq, by = c("lkey" = "key")) else mutate(camp, lq_lots = NA_integer_, lq_pass = NA_real_, lq_informed = NA_real_, lq_noncomp = NA_real_, lq_last = as.Date(NA))
  # an LQAS result that ended before the campaign started cannot belong to it (the same campaign code is often reused by another country)
  bad_lq <- !is.na(camp$lq_last) & camp$lq_last < camp$date - 14
  if (any(bad_lq)) { camp$lq_lots[bad_lq] <- NA; camp$lq_pass[bad_lq] <- NA; camp$lq_informed[bad_lq] <- NA; camp$lq_noncomp[bad_lq] <- NA; camp$lq_last[bad_lq] <- as.Date(NA); camp$lkey[bad_lq] <- NA_character_ }
  camp$lq_expected <- !is.na(lqas_max) & (camp$date + CFG$lqas_grace_days <= lqas_max)

  yn <- function(ok, na) ifelse(na, "N/E", ifelse(ok, "Yes", "No"))
  res <- camp %>% transmute(
    key, 
    COV1 = yn(!is.na(cov_total) & cov_total >= CFG$cov_std & cov_total <= CFG$cov_ceiling, is.na(cov_total)),
    COV2 = yn(!is.na(share_ge95) & share_ge95 >= CFG$share_ge95_std, is.na(share_ge95)),
    COV3 = yn(!is.na(share_low) & share_low <= CFG$share_low_max, is.na(share_low)),
    DQ1  = yn(complete >= CFG$complete_std, FALSE),
    DQ2  = yn(implaus <= CFG$implaus_max, FALSE),
    DQ3  = yn(!is.na(med_delay) & med_delay >= 0 & med_delay <= CFG$delay_max_days, report_share < CFG$report_min_share | is.na(med_delay)),
    DQ4  = yn(isTRUE(code_ok) | code_ok %in% TRUE, is.na(code_ok)),
    LQ1  = ifelse(!is.na(lq_lots), "Yes", ifelse(lq_expected, "No", "N/E")),
    LQ2  = yn(!is.na(lq_pass) & lq_pass >= CFG$lqas_pass_std, is.na(lq_pass)),
    LQ3  = yn(!is.na(lq_informed) & lq_informed >= CFG$informed_std, is.na(lq_informed)),
    LQ4  = yn(!is.na(lq_noncomp) & lq_noncomp <= CFG$noncomp_max, is.na(lq_noncomp)))
  vals <- camp %>% transmute(key, COV1 = cov_total, COV2 = share_ge95, COV3 = share_low, DQ1 = complete, DQ2 = implaus,
                             DQ3 = med_delay, DQ4 = as.numeric(code_ok), LQ1 = lq_lots, LQ2 = lq_pass, LQ3 = lq_informed, LQ4 = lq_noncomp)
  long <- res %>% pivot_longer(-key, names_to = "id", values_to = "result") %>%
    left_join(vals %>% pivot_longer(-key, names_to = "id", values_to = "value"), by = c("key", "id")) %>%
    left_join(IND, by = "id")
  sc <- long %>% group_by(key) %>% summarise(n_yes = sum(result == "Yes"), n_no = sum(result == "No"), n_ne = sum(result == "N/E"), .groups = "drop") %>%
    mutate(n_eval = n_yes + n_no, score_raw = ifelse(n_eval > 0, 100 * n_yes / n_eval, NA_real_))
  camp <- camp %>% left_join(sc, by = "key") %>%
    mutate(evaluable = !small & n_eval >= CFG$min_evaluable,
           score = ifelse(evaluable, round(score_raw), NA_real_),
           status = ifelse(small, "Not scored (small round)", ifelse(evaluable, "Evaluable", "N/E")),
           evidence = ifelse(!evaluable, NA, ifelse(n_eval >= CFG$full_evidence, "Full", ifelse(n_eval >= CFG$min_evaluable + 2, "Partial", "Limited"))),
           period = qlabel(date), period_key = qkey(date), month = format(date, "%b %Y"),
           country = nice_country(Country)) %>%
    arrange(Country, date)
  long <- long %>% left_join(camp %>% select(key, country, Response, Round_Add, date, period, period_key, evaluable, small), by = "key")
  list(camp = camp, long = long)
}

country_table <- function(camp, long) {
  sc <- camp %>% filter(!small)
  out <- sc %>% group_by(Country, country) %>% group_modify(function(d, k) {
    d <- arrange(d, date); e <- filter(d, evaluable)
    last_any <- d[nrow(d), ]
    if (!nrow(e)) return(tibble(n_campaigns = nrow(d), n_eval = 0L, first_score = NA_real_, latest_score = NA_real_, latest_response = last_any$Response,
      latest_round = last_any$Round_Add, latest_date = last_any$date, first_date = NA, latest_is_last = FALSE, latest_evidence = NA_character_, latest_implaus = last_any$implaus,
      failed = ""))
    l <- e[nrow(e), ]; f <- e[1, ]
    fl <- long %>% filter(key == l$key, result == "No") %>% pull(label)
    tibble(n_campaigns = nrow(d), n_eval = nrow(e), first_score = f$score, latest_score = l$score, latest_response = l$Response,
           latest_round = l$Round_Add, latest_date = l$date, first_date = f$date, latest_is_last = identical(l$key, last_any$key),
           latest_evidence = l$evidence, latest_implaus = l$implaus, failed = paste(fl, collapse = "; "))
  }) %>% ungroup()
  out <- out %>% mutate(
    change = ifelse(n_eval >= 2, latest_score - first_score, NA_real_),
    trend = case_when(n_eval == 0 ~ "No evaluable campaign", n_eval == 1 ~ "Single evaluable campaign",
                      change > CFG$trend_tol ~ "Improving", change < -CFG$trend_tol ~ "Declining", TRUE ~ "No change"),
    band = band_of(latest_score), risk = risk_of(band),
    segment = case_when(n_eval == 0 ~ "Urgent validation",
                        latest_score < CFG$bands[1] | (latest_score < CFG$bands[2] & trend == "Declining") ~ "Reinforced support",
                        latest_score < CFG$bands[2] ~ "Close monitoring",
                        latest_score < CFG$bands[3] ~ "Targeted follow-up", TRUE ~ "Consolidation"),
    needs_validation = n_eval == 0 | !latest_is_last | latest_evidence %in% "Limited" | coalesce(latest_implaus, 0) > 25)
  out
}

period_tables <- function(camp, long) {
  pk <- camp %>% filter(evaluable) %>% distinct(period, period_key) %>% arrange(period_key)
  keep <- tail(pk, 2)
  ev <- long %>% filter(evaluable, period %in% keep$period)
  ind <- ev %>% filter(result != "N/E") %>% group_by(id, pillar, label, period, period_key) %>%
    summarise(n = n(), met = sum(result == "Yes"), rate = 100 * met / n, .groups = "drop")
  pil <- ev %>% filter(result != "N/E") %>% group_by(pillar, period, period_key) %>%
    summarise(n = n(), met = sum(result == "Yes"), rate = 100 * met / n, .groups = "drop")
  list(periods = keep, ind = ind, pillar = pil)
}
# LQAS passed, by quarter: all campaigns and a zoom on selected countries.
# The two quarters shown are always the last two quarters that contain campaigns in the
# review window, even when the LQAS file has no result for them (the chart then says so).
lqas_tables <- function(lqas, camp, zoom) {
  per <- camp %>% distinct(period, period_key) %>% arrange(period_key) %>% tail(2)
  periods <- per$period
  grid_all <- tibble(period = per$period, period_key = per$period_key)
  lots <- tibble(key = character(), country = character(), period = character(), period_key = numeric(), pass = logical())
  if (!is.null(lqas) && nrow(lqas)) {
    cm <- camp %>% filter(!is.na(lkey), !is.na(lq_lots)) %>% distinct(lkey, .keep_all = TRUE) %>% select(lkey, country, period, period_key)
    lots <- lqas %>% inner_join(cm, by = c("key" = "lkey")) %>% filter(period %in% periods)
  }
  cam <- lots %>% group_by(key, country, period, period_key) %>% summarise(lots = n(), passed = sum(pass), .groups = "drop") %>%
    mutate(accepted = 100 * passed / lots >= CFG$lqas_pass_std)
  all <- lots %>% group_by(period, period_key) %>% summarise(lots = n(), passed = sum(pass), rate = 100 * passed / lots, .groups = "drop") %>%
    left_join(cam %>% group_by(period) %>% summarise(campaigns = n(), accepted = sum(accepted), accepted_rate = 100 * accepted / campaigns), by = "period")
  all <- grid_all %>% left_join(all %>% select(-period_key), by = "period") %>%
    mutate(across(c(lots, passed, campaigns, accepted), ~ ifelse(is.na(.x), 0L, as.integer(.x)))) %>% arrange(period_key)
  zm <- lots %>% filter(country %in% zoom) %>% group_by(country, period, period_key) %>%
    summarise(lots = n(), passed = sum(pass), rate = 100 * passed / lots, campaigns = n_distinct(key), .groups = "drop") %>% arrange(country, period_key)
  if (!nrow(lots)) message("  ! No LQAS result matches the campaigns of ", paste(periods, collapse = " / "), " - the LQAS slides will say so. Provide a more recent LQAS file with --lqas.")
  else for (z in setdiff(zoom, zm$country)) message("  ! No LQAS result for ", z, " in ", paste(periods, collapse = " / "))
  list(all = all, zoom = zm, periods = periods)
}

ch_lqas_all <- function(lt) {
  d <- lt$all; if (!length(lt$periods)) return(NULL)
  d <- d %>% mutate(period = factor(period, levels = lt$periods), none = lots == 0)
  long <- bind_rows(
    tibble(period = d$period, measure = "LQAS lots passed", rate = ifelse(d$none, 0, d$rate), none = d$none,
           lab = ifelse(d$none, "no LQAS result", sprintf("%d%%\n(%s of %s lots)", round(d$rate), comma(d$passed), comma(d$lots)))),
    tibble(period = d$period, measure = sprintf("Campaigns accepted (\u2265%d%% lots passed)", CFG$lqas_pass_std), rate = ifelse(d$none, 0, d$accepted_rate), none = d$none,
           lab = ifelse(d$none, "no LQAS result", sprintf("%d%%\n(%d of %d campaigns)", round(d$accepted_rate), d$accepted, d$campaigns))))
  cols <- setNames(c(COL$grey, COL$navy)[seq_along(lt$periods) + (2 - length(lt$periods))], lt$periods)
  ggplot(long, aes(period, rate, fill = period)) + geom_col(width = 0.62) +
    geom_text(aes(label = lab, colour = none), vjust = -0.3, size = 3.6, fontface = "bold", lineheight = 0.95) +
    scale_colour_manual(values = c(`FALSE` = COL$text, `TRUE` = COL$muted), guide = "none") +
    facet_wrap(~ measure) + scale_fill_manual(values = cols, guide = "none") +
    scale_y_continuous(limits = c(0, 118), breaks = seq(0, 100, 25), labels = function(x) paste0(x, "%"), expand = c(0, 0)) +
    labs(title = "Proportion of LQAS passed, all campaigns") + theme_sia() +
    theme(strip.text = element_text(face = "bold", colour = COL$navy), panel.grid.major.y = element_line(colour = "#ECEFF3"), panel.grid.major.x = element_blank())
}

ch_lqas_zoom <- function(lt, zoom) {
  d <- lt$zoom; if (!length(lt$periods)) return(NULL)
  grid <- expand.grid(country = zoom, period = lt$periods, stringsAsFactors = FALSE) %>% as_tibble() %>% left_join(d, by = c("country", "period"))
  grid <- grid %>% mutate(period = factor(period, levels = lt$periods), country = factor(country, levels = zoom),
                          lab = ifelse(is.na(rate), "no\nLQAS", sprintf("%d%%\n(%d lots)", round(rate), lots)), y = ifelse(is.na(rate), 0, rate))
  cols <- setNames(c(COL$grey, COL$navy)[seq_along(lt$periods) + (2 - length(lt$periods))], lt$periods)
  ggplot(grid, aes(country, y, fill = period)) + geom_col(position = position_dodge2(width = 0.8), width = 0.75) +
    geom_text(aes(label = lab, colour = is.na(rate)), position = position_dodge2(width = 0.8), vjust = -0.3, size = 3.2, lineheight = 0.95, show.legend = FALSE) +
    scale_colour_manual(values = c(`FALSE` = COL$text, `TRUE` = COL$muted)) + scale_fill_manual(values = cols) +
    scale_y_continuous(limits = c(0, 122), breaks = seq(0, 100, 25), labels = function(x) paste0(x, "%"), expand = c(0, 0)) +
    labs(title = sprintf("Proportion of LQAS lots passed: %s", paste(zoom, collapse = ", "))) + theme_sia() +
    theme(panel.grid.major.y = element_line(colour = "#ECEFF3"), panel.grid.major.x = element_blank())
}

# ----------------------------------------------------------------------------- 6. CHARTS
theme_sia <- function(base = 12) {
  theme_minimal(base_size = base, base_family = "sans") +
    theme(plot.title = element_text(face = "bold", colour = COL$navy, size = base + 1, hjust = 0.5),
          plot.subtitle = element_text(colour = COL$muted, hjust = 0.5),
          plot.caption = element_text(colour = COL$muted, size = base - 3, hjust = 0),
          axis.text = element_text(colour = COL$text), axis.title = element_blank(),
          panel.grid.major.y = element_blank(), panel.grid.minor = element_blank(),
          panel.grid.major.x = element_line(colour = "#ECEFF3"), legend.position = "top",
          legend.title = element_blank(), plot.margin = margin(8, 14, 8, 8))
}
band_fill <- function(b) setNames(c(COL$red, COL$orange, COL$gold, COL$green), levels(b))

ch_ranking <- function(ct, ttl) {
  d <- ct %>% filter(!is.na(latest_score)) %>% mutate(country = fct_reorder(country, latest_score))
  ggplot(d, aes(latest_score, country, fill = band)) + geom_col(width = 0.72) +
    geom_text(aes(label = paste0(latest_score, "%")), hjust = -0.15, fontface = "bold", colour = COL$text, size = 3.6) +
    scale_fill_manual(values = band_fill(d$band), guide = "none") +
    scale_x_continuous(limits = c(0, 108), breaks = seq(0, 100, 20), labels = function(x) paste0(x, "%"), expand = c(0, 0)) +
    labs(title = ttl) + theme_sia()
}

ch_trend <- function(ct) {
  d <- ct %>% filter(!is.na(change)) %>% mutate(country = fct_reorder(country, change), dir = change >= 0)
  if (!nrow(d)) return(NULL)
  lim <- max(abs(d$change), 10) * 1.2
  ggplot(d, aes(change, country, fill = dir)) + geom_col(width = 0.72) +
    geom_text(aes(label = sprintf("%+d", as.integer(change)), hjust = ifelse(dir, -0.2, 1.2)), fontface = "bold", colour = COL$text, size = 3.6) +
    geom_vline(xintercept = 0, colour = "#9CA3AF") +
    scale_fill_manual(values = c(`TRUE` = COL$green, `FALSE` = COL$red), guide = "none") +
    scale_x_continuous(limits = c(-lim, lim)) +
    labs(title = "Change from first to latest evaluable campaign (percentage points)") + theme_sia()
}

ch_pillar <- function(pt) {
  if (is.null(pt) || !nrow(pt$ind)) return(NULL)
  lv <- pt$periods$period
  d <- pt$ind %>% mutate(period = factor(period, levels = lv),
                         label = factor(label, levels = rev(unique(IND$label[IND$label %in% label]))))
  cols <- setNames(c(COL$grey, COL$navy)[seq_len(length(lv)) + (2 - length(lv))], lv)
  ggplot(d, aes(rate, label, fill = period)) +
    geom_col(position = position_dodge2(reverse = TRUE, width = 0.8), width = 0.75) +
    geom_text(aes(label = paste0(round(rate), "%")), position = position_dodge2(reverse = TRUE, width = 0.8), hjust = -0.15, size = 3.2, colour = COL$text) +
    scale_fill_manual(values = cols, breaks = lv) +
    scale_x_continuous(limits = c(0, 112), breaks = seq(0, 100, 25), labels = function(x) paste0(x, "%"), expand = c(0, 0)) +
    labs(title = "Compliance by indicator (% of evaluable campaigns at standard)") + theme_sia(11)
}

ch_heat <- function(ct, long) {
  keys <- ct %>% filter(!is.na(latest_score)) %>% transmute(country, key = paste(Country, latest_response, latest_round, sep = "|"), latest_score)
  d <- long %>% inner_join(keys, by = c("key", "country")) %>%
    mutate(country = fct_reorder(country, latest_score), label = factor(label, levels = IND$label),
           result = factor(result, levels = c("Yes", "No", "N/E")))
  ggplot(d, aes(label, country, fill = result)) + geom_tile(colour = "white", linewidth = 0.8) +
    scale_fill_manual(values = c(Yes = COL$green, No = COL$red, `N/E` = "#D1D5DB"), drop = FALSE) +
    scale_x_discrete(position = "top", labels = function(x) str_wrap(x, 12)) +
    labs(title = "Indicators met in each country's latest evaluable campaign") +
    theme_sia(10) + theme(axis.text.x = element_text(size = 8), panel.grid = element_blank())
}

ch_dq <- function(camp, long) {
  d <- long %>% filter(evaluable, pillar == "Data quality", result != "N/E") %>% group_by(label, period, period_key) %>%
    summarise(rate = 100 * mean(result == "No"), .groups = "drop")
  if (!nrow(d)) return(NULL)
  lv <- unique(arrange(d, period_key)$period)
  d <- d %>% mutate(period = factor(period, levels = lv))
  ggplot(d, aes(rate, fct_rev(factor(label, levels = IND$label)), fill = period)) +
    geom_col(position = position_dodge2(reverse = TRUE), width = 0.7) +
    geom_text(aes(label = paste0(round(rate), "%")), position = position_dodge2(reverse = TRUE, width = 0.7), hjust = -0.2, size = 3.3) +
    scale_fill_manual(values = setNames(c(COL$grey, COL$red, COL$navy, COL$orange)[seq_along(lv)], lv)) +
    scale_x_continuous(limits = c(0, 100), labels = function(x) paste0(x, "%"), expand = expansion(mult = c(0, 0.08))) +
    labs(title = "Data-quality indicators NOT met (% of evaluable campaigns)") + theme_sia()
}

save_png <- function(p, file, w = 9, h = 6.2) {
  if (is.null(p)) return(invisible(NULL))
  ragg::agg_png(file, width = w, height = h, units = "in", res = 200, background = "white")
  print(p); dev.off(); invisible(file)
}

# ----------------------------------------------------------------------------- 7. KEY NUMBERS, MESSAGES, DECISIONS
ACTION_BY_PILLAR <- c(Coverage = "close the coverage gaps with targeted revisits, supervision and social mobilisation",
                      `Data quality` = "validate and complete the district data before the next round",
                      `LQAS verification` = "run LQAS on time, revaccinate failed lots and document closure")

list_names <- function(x, max = 12) {
  x <- as.character(x)
  if (!length(x)) return("none")
  if (length(x) > max) paste0(paste(x[seq_len(max)], collapse = " • "), " • +", length(x) - max)
  else paste(x, collapse = " • ")
}

make_key_numbers <- function(camp, units, ct, pt, from, to, lqas_max, long) {
  sc <- camp %>% filter(!small); ev <- sc %>% filter(evaluable)
  per <- pt$periods$period
  mean_by <- ev %>% filter(period %in% per) %>% group_by(period, period_key) %>%
    summarise(mean = mean(score), n = n(), .groups = "drop") %>% arrange(period_key)
  ne_by <- sc %>% filter(period %in% per) %>% group_by(period, period_key) %>%
    summarise(campaigns = n(), ne = sum(!evaluable), .groups = "drop") %>% arrange(period_key)
  top <- ct %>% filter(!is.na(latest_score)) %>% arrange(desc(latest_score))
  list(from = from, to = to, lqas_max = lqas_max,
       n_campaigns = nrow(sc), n_eval = nrow(ev), n_ne = sum(!sc$evaluable), n_small = sum(camp$small),
       n_countries = nrow(ct), n_countries_scored = nrow(top),
       mean_score = mean(ev$score), mean_by = mean_by, ne_by = ne_by,
       n_below = sum(top$latest_score < CFG$bands[2]), n_vlow = sum(top$latest_score < CFG$bands[1]),
       best = top[1, ], worst = top[nrow(top), ],
       gaps = sum(long$result == "No" & long$evaluable), n_districts = nrow(units),
       n_missing = sum(units$missing), n_implaus = sum(units$implausible), n_low = sum(units$low, na.rm = TRUE),
       n_code_bad = sum(sc$code_ok %in% FALSE), n_limited = sum(ev$evidence %in% "Limited"))
}

make_messages <- function(K, ct, pt) {
  per <- pt$periods$period; mb <- K$mean_by
  m1 <- if (nrow(mb) >= 2) {
    d <- mb$mean[2] - mb$mean[1]
    sprintf("The mean campaign score %s from %s in %s to %s in %s (%+.1f points), but the %s estimate rests on %d evaluable campaign%s and the country mix differs: treat it as indicative.",
            ifelse(d >= 0, "rose", "fell"), pct(mb$mean[1], 1), mb$period[1], pct(mb$mean[2], 1), mb$period[2], d, mb$period[2], mb$n[2], ifelse(mb$n[2] == 1, "", "s"))
  } else sprintf("The mean campaign score is %s across %d evaluable campaigns; one quarter only, so no comparison over time is possible.", pct(K$mean_score, 1), K$n_eval)
  imp <- ct %>% filter(trend == "Improving") %>% arrange(desc(change)); dec <- ct %>% filter(trend == "Declining") %>% arrange(change)
  m2 <- sprintf("%s %s. %s",
                ifelse(nrow(imp), sprintf("Gains in %s", paste0(head(imp$country, 2), " (", sprintf("%+d", as.integer(head(imp$change, 2))), ")", collapse = " and ")), "No country improved"),
                ifelse(nrow(dec), sprintf("while %s declined (%s)", paste0(head(dec$country, 2), collapse = " and "), paste(sprintf("%+d", as.integer(head(dec$change, 2))), collapse = ", ")), "and no country declined"),
                sprintf("%d of %d countries have a latest score below %d%%.", K$n_below, K$n_countries_scored, CFG$bands[2]))
  m3 <- sprintf("%s district records lack a target or vaccinated figure, %s show implausible coverage (above %d%%), %d campaign%s %s not evaluable and %d campaign code%s do not match the country.",
                comma(K$n_missing), comma(K$n_implaus), CFG$cov_ceiling, K$n_ne, ifelse(K$n_ne == 1, "", "s"), ifelse(K$n_ne == 1, "is", "are"), K$n_code_bad, ifelse(K$n_code_bad == 1, "", "s"))
  weak <- pt$ind %>% filter(period == tail(per, 1)) %>% arrange(rate) %>% head(3)
  m4 <- sprintf("The weakest indicators in %s are %s: the response must target these causes, not only the scores.",
                tail(per, 1), paste0(weak$label, " (", round(weak$rate), "%)", collapse = ", "))
  list(
    list(title = if (nrow(mb) >= 2 && mb$mean[2] >= mb$mean[1]) "Progress exists, but is it widespread?" else if (nrow(mb) >= 2) "Scores slipped between quarters" else "One quarter only: no trend yet", body = m1),
    list(title = "Country trajectories diverge", body = m2),
    list(title = "Data quality may mask actual performance", body = m3),
    list(title = "The response must target causes, not only scores", body = m4))
}

make_decisions <- function(ct, K) {
  below <- ct %>% filter(!is.na(latest_score), latest_score < CFG$bands[2]) %>% pull(country)
  val <- ct %>% filter(needs_validation) %>% pull(country)
  tibble(
    Action = c(paste0("Validate ", ifelse(K$n_ne > 0, sprintf("the %d campaign%s not evaluable and ", K$n_ne, ifelse(K$n_ne == 1, "", "s")), ""), sprintf("the %s implausible or incomplete district records (%s)", comma(K$n_missing + K$n_implaus), ifelse(length(val), list_names(val, 6), "no country flagged"))),
               sprintf("Develop corrective action plans for the %d countries scoring below %d%% (%s)", length(below), CFG$bands[2], list_names(below, 6)),
               "Review critical gaps monthly with the same score and indicators",
               "Verify closure of post-LQAS action points, including failure analysis and revaccination"),
    Responsible = c("Countries", "Countries + AFRO", "AFRO", "Countries"),
    Timeline = c(CFG$days_validate, CFG$days_plan, "Monthly", "Before the following round"),
    Evidence = c("Updated dataset and validation note", "Plan with pillar, owner, deadline and evidence", "Regional tracker and management review", "Closure evidence attached to the next review"))
}

segment_table <- function(ct) {
  defs <- tribble(~segment, ~colour, ~action,
    "Reinforced support", COL$red, sprintf("Corrective plan within %s", CFG$days_plan),
    "Close monitoring", COL$orange, "Monthly review and follow-up of priority gaps",
    "Targeted follow-up", COL$navy, "Correct specific indicators and verify sustainability",
    "Consolidation", COL$green, "Sustain gains and document post-LQAS action",
    "Urgent validation", COL$blue, sprintf("Complete or justify the data within %s", CFG$days_validate))
  defs %>% rowwise() %>% mutate(
    countries = list(if (segment == "Urgent validation") ct$country[ct$needs_validation] else ct$country[ct$segment == segment]),
    n = length(countries)) %>% ungroup()
}

# ----------------------------------------------------------------------------- 8. EXCEL
write_workbook <- function(res, file) {
  library(openxlsx)
  wb <- createWorkbook(); hs <- createStyle(fontColour = "white", fgFill = COL$navy, textDecoration = "bold", halign = "center", valign = "center", wrapText = TRUE, border = "bottom")
  add <- function(name, df, widths = NULL) {
    addWorksheet(wb, name); writeData(wb, name, df, headerStyle = hs, withFilter = TRUE); freezePane(wb, name, firstRow = TRUE)
    setColWidths(wb, name, cols = seq_len(ncol(df)), widths = if (is.null(widths)) "auto" else widths)
  }
  K <- res$K
  readme <- tibble(Item = c("Title", "Period", "Data (admin)", "Data (LQAS)", "LQAS data end", "Campaign rounds", "Evaluable", "Not evaluable (N/E)", "Rounds too small to score", "Score", "Indicators", "Band", "Trend", "Segments", "Limits"),
    Description = c(res$title, paste(format(K$from), "to", format(K$to)), res$admin_file, res$lqas_file %||% "not found", ifelse(is.na(K$lqas_max), "n/a", format(K$lqas_max)),
      K$n_campaigns, K$n_eval, K$n_ne, K$n_small,
      "Score = indicators met / indicators evaluable x 100 (equal weights). A campaign with fewer than the minimum number of evaluable indicators is N/E and is not scored.",
      "See sheet Indicators: pillar, standard, and the number met / not met / N/E for every campaign in Indicator_results.",
      sprintf("Latest evaluable score: below %d%% | %d-%d%% | %d-%d%% | %d%% or above", CFG$bands[1], CFG$bands[1], CFG$bands[2] - 1, CFG$bands[2], CFG$bands[3] - 1, CFG$bands[3]),
      sprintf("First vs latest evaluable campaign; change within +/-%d points = No change", CFG$trend_tol),
      "Reinforced support: score below the first band or below the second band and declining | Close monitoring | Targeted follow-up | Consolidation | Urgent validation (cross-cutting: N/E, partial evidence, or implausible data)",
      "The standards in sheet Standards are working defaults to be confirmed by the technical team. Admin coverage is not an independent verification; LQAS is used where it exists."))
  add("README", readme, c(26, 120))
  add("Country_summary", res$country %>% transmute(Country = country, `Latest score` = latest_score, Band = as.character(band), Risk = as.character(risk), Trend = trend, `First score` = first_score, Change = change,
        `Latest campaign` = latest_response, Round = latest_round, `Latest date` = latest_date, `Evaluable campaigns` = n_eval, `Campaigns in period` = n_campaigns,
        Evidence = latest_evidence, Segment = segment, `Needs validation` = ifelse(needs_validation, "Yes", ""), `Indicators not met (latest)` = failed), c(22, 10, 14, 11, 24, 10, 9, 36, 8, 12, 11, 11, 10, 20, 11, 70))
  add("Campaign_scores", res$camp %>% transmute(Country = country, Response, Round = Round_Add, Date = date, Period = period, Vaccine = vaccine, Districts = n_districts, Status = status, Score = score,
        `Indicators met` = n_yes, `Not met` = n_no, `N/E` = n_ne, Evidence = evidence, `Coverage %` = round(cov_total, 1), `Districts >= std %` = round(share_ge95, 1),
        `Districts < low %` = round(share_low, 1), `Complete %` = round(complete, 1), `Implausible %` = round(implaus, 1), `Median report delay (d)` = med_delay,
        `LQAS lots` = lq_lots, `LQAS pass %` = round(lq_pass, 1)), c(20, 34, 8, 12, 10, 16, 10, 22, 8, 9, 8, 6, 9, 10, 12, 12, 10, 11, 12, 9, 10))
  add("Indicator_results", res$long %>% transmute(Country = country, Response, Round = Round_Add, Date = date, Period = period, Pillar = pillar, Indicator = label, Standard = standard_text, Value = round(value, 1), Result = result), c(20, 34, 8, 12, 10, 18, 28, 60, 10, 9))
  add("Indicators", IND %>% transmute(ID = id, Pillar = pillar, Indicator = label, Standard = standard_text), c(8, 20, 30, 70))
  add("By_indicator_period", res$periods$ind %>% transmute(ID = id, Pillar = pillar, Indicator = label, Period = period, `Campaigns evaluated` = n, Met = met, `Share met %` = round(rate, 1)), c(8, 20, 30, 10, 12, 8, 12))
  add("By_pillar_period", res$periods$pillar %>% transmute(Pillar = pillar, Period = period, `Indicator results` = n, Met = met, `Share met %` = round(rate, 1)), c(22, 10, 14, 8, 12))
  if (nrow(res$lqas_t$all)) {
    add("LQAS_passed_all", res$lqas_t$all %>% transmute(Period = period, `LQAS lots` = lots, `Lots passed` = passed, `Lots passed %` = round(rate, 1), Campaigns = campaigns, `Campaigns accepted` = accepted, `Campaigns accepted %` = round(accepted_rate, 1)), c(12, 12, 12, 14, 12, 18, 20))
    add("LQAS_passed_zoom", res$lqas_t$zoom %>% transmute(Country = country, Period = period, `LQAS lots` = lots, `Lots passed` = passed, `Lots passed %` = round(rate, 1), Campaigns = campaigns), c(20, 12, 12, 12, 14, 12))
  }
  add("Segments", res$segments %>% transmute(Segment = segment, Countries = n, List = map_chr(countries, ~ paste(.x, collapse = ", ")), Action = action), c(22, 10, 80, 60))
  dq <- tibble(Measure = c("District records", "Without target or vaccinated figure", sprintf("Coverage above %d%%", CFG$cov_ceiling), sprintf("Coverage below %d%%", CFG$cov_low), "Campaign rounds scored", "Not evaluable (N/E)", "Evaluable with limited evidence", "Campaign code not matching the country", "Rounds too small to score", "Indicator results not met (gaps)"),
               Value = c(K$n_districts, K$n_missing, K$n_implaus, K$n_low, K$n_campaigns, K$n_ne, K$n_limited, K$n_code_bad, K$n_small, K$gaps))
  add("Data_quality", dq, c(46, 14))
  add("District_flags", res$units %>% filter(missing | implausible) %>% transmute(Country = nice_country(Country), Response, Round = Round_Add, Province, District, Date = sia_date, Target = pop, Vaccinated = vac, `Coverage %` = round(cv, 1),
        Flag = ifelse(missing, "Missing target or vaccinated", "Implausible coverage")) %>% head(20000), c(20, 34, 8, 22, 26, 12, 12, 12, 11, 28))
  add("Decisions", res$decisions, c(90, 18, 40, 48))
  if (!is.null(res$lqas_unmatched) && nrow(res$lqas_unmatched)) add("LQAS_not_matched", res$lqas_unmatched)
  add("Standards", tibble(Setting = names(CFG)[map_lgl(CFG, ~ is.numeric(.x) && length(.x) == 1)], Value = unlist(CFG[map_lgl(CFG, ~ is.numeric(.x) && length(.x) == 1)])), c(24, 12))
  cs <- createStyle(bgFill = "#F8D7DA"); cg <- createStyle(bgFill = "#D8EBD9")
  for (sh in c("Country_summary", "Campaign_scores")) {
    col <- if (sh == "Country_summary") 2 else 9
    conditionalFormatting(wb, sh, cols = col, rows = 2:2000, rule = paste0("AND(ISNUMBER(", int2col(col), "2),", int2col(col), "2<", CFG$bands[2], ")"), style = cs)
    conditionalFormatting(wb, sh, cols = col, rows = 2:2000, rule = paste0("AND(ISNUMBER(", int2col(col), "2),", int2col(col), "2>=", CFG$bands[3], ")"), style = cg)
  }
  saveWorkbook(wb, file, overwrite = TRUE)
}

# ----------------------------------------------------------------------------- 9. NARRATIVE (HTML + Markdown)
write_narrative <- function(res, html_file, md_file) {
  K <- res$K; ct <- res$country; msgs <- res$messages
  span <- sprintf("%s – %s", format(K$from, "%b %Y"), format(K$to, "%b %Y"))
  esc <- function(x) htmltools::htmlEscape(x)
  tile <- function(big, label, sub, col = COL$navy) sprintf('<div style="flex:1;min-width:150px;border:1px solid #D9E0E8;border-radius:10px;padding:12px 16px;background:#fff"><div style="font-size:28px;font-weight:800;color:%s">%s</div><div style="font-weight:700;color:#1F2937">%s</div><div style="color:#6B7280;font-size:13px">%s</div></div>', col, esc(big), esc(label), esc(sub))
  mb <- K$mean_by
  tiles <- paste0(
    tile(pct(K$mean_score, 1), "Mean campaign score", sprintf("%d evaluable campaigns", K$n_eval)),
    if (nrow(mb) >= 2) tile(pct(mb$mean[2], 1), sprintf("%s mean score", mb$period[2]), sprintf("vs %s in %s", pct(mb$mean[1], 1), mb$period[1])) else "",
    tile(K$n_campaigns, "Campaign rounds scored", sprintf("%d countries", K$n_countries)),
    tile(K$n_ne, "Campaigns not evaluable", ifelse(K$n_ne == 0, "all evaluable", "to complete or justify"), COL$red))
  cards <- paste0(map_chr(seq_along(msgs), function(i) sprintf('<div style="flex:1 1 45%%;min-width:260px;background:#EEF2F6;border:1px solid #D9E0E8;border-radius:10px;padding:12px 16px"><div style="font-weight:800;color:%s">%d &nbsp; %s</div><div style="color:#1F2937;margin-top:4px">%s</div></div>', COL$navy, i, esc(msgs[[i]]$title), esc(msgs[[i]]$body))), collapse = "")
  rows <- paste0(map_chr(seq_len(nrow(ct %>% filter(!is.na(latest_score)))), function(i) {
    r <- ct %>% filter(!is.na(latest_score)) %>% arrange(desc(latest_score)) %>% slice(i)
    sprintf('<tr><td>%s</td><td style="text-align:right;font-weight:700;color:%s">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>', esc(r$country), band_colour(r$band), paste0(r$latest_score, "%"), esc(r$trend), esc(r$segment), esc(r$failed))
  }), collapse = "")
  dec <- paste0(map_chr(seq_len(nrow(res$decisions)), function(i) with(res$decisions[i, ], sprintf("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>", esc(Action), esc(Responsible), esc(Timeline), esc(Evidence)))), collapse = "")
  html <- sprintf('<div style="font-family:Segoe UI,Arial,sans-serif;color:#1F2937;max-width:1100px">
<h2 style="color:%s;margin:0">%s</h2><div style="color:#6B7280;margin-bottom:12px">%s | %s | %d countries | %d campaign rounds</div>
<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px">%s</div>
<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px">%s</div>
<h3 style="color:%s">Latest evaluable campaign by country</h3>
<table style="border-collapse:collapse;width:100%%" border="0" cellpadding="5"><tr style="background:%s;color:#fff"><th align="left">Country</th><th align="right">Score</th><th align="left">Trend</th><th align="left">Segment</th><th align="left">Indicators not met</th></tr>%s</table>
<h3 style="color:%s">Decisions</h3><table style="border-collapse:collapse;width:100%%" border="0" cellpadding="5"><tr style="background:%s;color:#fff"><th align="left">Action</th><th align="left">Responsible</th><th align="left">Timeline</th><th align="left">Evidence</th></tr>%s</table>
<p style="color:#6B7280;font-size:12px">Score = indicators met / indicators evaluable x 100. Standards are working defaults (see the Standards sheet of the workbook). Source: AFRO administrative data%s.</p></div>',
    COL$navy, esc(res$title), esc(CFG$org), esc(span), K$n_countries, K$n_campaigns, tiles, cards, COL$navy, COL$navy, rows, COL$navy, COL$navy, dec,
    ifelse(is.null(res$lqas_file), "", " and LQAS data"))
  writeLines(html, html_file, useBytes = TRUE)
  md <- c(sprintf("# %s", res$title), sprintf("*%s | %s*", CFG$org, span), "",
          sprintf("**Mean campaign score %s** across %d evaluable campaigns; %d campaign(s) not evaluable; %d of %d countries below %d%%.", pct(K$mean_score, 1), K$n_eval, K$n_ne, K$n_below, K$n_countries_scored, CFG$bands[2]), "",
          unlist(map(seq_along(msgs), function(i) c(sprintf("**%d. %s**", i, msgs[[i]]$title), msgs[[i]]$body, ""))),
          "## Latest evaluable campaign by country", "", "| Country | Score | Trend | Segment | Indicators not met |", "|---|---:|---|---|---|",
          map_chr(seq_len(nrow(ct %>% filter(!is.na(latest_score)))), function(i) { r <- ct %>% filter(!is.na(latest_score)) %>% arrange(desc(latest_score)) %>% slice(i); sprintf("| %s | %s%% | %s | %s | %s |", r$country, r$latest_score, r$trend, r$segment, r$failed) }),
          "", "## Decisions", "", "| Action | Responsible | Timeline | Evidence |", "|---|---|---|---|",
          map_chr(seq_len(nrow(res$decisions)), function(i) with(res$decisions[i, ], sprintf("| %s | %s | %s | %s |", Action, Responsible, Timeline, Evidence))))
  writeLines(md, md_file, useBytes = TRUE)
}

# ----------------------------------------------------------------------------- 10. POWERPOINT
FONT <- "Arial"
fpt <- function(size = 14, bold = FALSE, col = COL$text, italic = FALSE) officer::fp_text(font.size = size, bold = bold, color = col, italic = italic, font.family = FONT)
ch  <- function(txt, size = 14, bold = FALSE, col = COL$text, italic = FALSE) flextable::as_chunk(txt, props = fpt(size, bold, col, italic))
brd <- function(col = "#D9E0E8", w = 1) officer::fp_border(color = col, width = w)

make_template_169 <- function() {
  tpl <- system.file("template", "template.pptx", package = "officer")
  tmp <- tempfile("tpl"); dir.create(tmp); utils::unzip(tpl, exdir = tmp)
  f <- file.path(tmp, "ppt", "presentation.xml")
  x <- paste(readLines(f, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  x <- sub("<p:sldSz[^>]*/>", '<p:sldSz cx="12192000" cy="6858000"/>', x)
  writeLines(x, f, useBytes = TRUE)
  out <- tempfile(fileext = ".pptx"); old <- setwd(tmp); on.exit(setwd(old))
  zip::zip(out, files = list.files(".", recursive = TRUE, all.files = TRUE), mode = "mirror")
  out
}

tbl_base <- function(ft, w, h = NULL) {
  ft <- ft %>% flextable::border_remove() %>% flextable::padding(padding.left = 8, padding.right = 8, padding.top = 4, padding.bottom = 4) %>%
    flextable::set_table_properties(layout = "fixed")
  ft <- flextable::width(ft, width = w / flextable::ncol_keys(ft)) 
  ft
}

ft_tile <- function(big, big_col, label, sub, w, h) {
  ft <- flextable::flextable(data.frame(a = ""))
  ft <- flextable::delete_part(ft, "header")
  ft <- flextable::compose(ft, j = 1, value = flextable::as_paragraph(ch(big, 30, TRUE, big_col), ch("\n"), ch(label, 13, TRUE, COL$text), ch(if (nzchar(sub)) paste0("\n", sub) else "", 11, FALSE, COL$muted)))
  ft <- tbl_base(ft, w) %>% flextable::border_outer(border = brd()) %>% flextable::bg(bg = "white") %>% flextable::valign(valign = "center")
  ft <- flextable::height(ft, height = h) %>% flextable::hrule(rule = "exact") %>% flextable::width(width = w)
  ft
}

ft_banner <- function(txt, w, h = 0.55, bg = COL$navy, col = "white", size = 16) {
  ft <- flextable::flextable(data.frame(a = "")) %>% flextable::delete_part("header")
  ft <- flextable::compose(ft, j = 1, value = flextable::as_paragraph(ch(txt, size, TRUE, col))) %>% tbl_base(w) %>%
    flextable::bg(bg = bg) %>% flextable::align(align = "center") %>% flextable::valign(valign = "center")
  flextable::height(ft, height = h) %>% flextable::hrule(rule = "exact") %>% flextable::width(width = w)
}

# lines: list of list(label, value, colour)
ft_panel <- function(header, lines, col, w, h, hh = 0.4, size = 13) {
  ft <- flextable::flextable(data.frame(a = c(header, ""))) %>% flextable::delete_part("header")
  ft <- flextable::compose(ft, i = 1, j = 1, value = flextable::as_paragraph(ch(header, 13, TRUE, "white")))
  body <- if (!length(lines)) list(ch("Nothing to report", size, FALSE, COL$muted)) else
    unlist(lapply(seq_along(lines), function(i) { L <- lines[[i]]
      c(list(ch(if (i > 1) "\n" else "", size), ch(paste0("• ", L$label), size, FALSE, COL$text), ch(L$value, size, TRUE, L$col %||% col))) }), recursive = FALSE)
  ft <- flextable::compose(ft, i = 2, j = 1, value = do.call(flextable::as_paragraph, body))
  ft <- tbl_base(ft, w) %>% flextable::bg(i = 1, bg = col) %>% flextable::bg(i = 2, bg = "white") %>% flextable::align(i = 1, align = "center") %>%
    flextable::valign(valign = "center") %>% flextable::border_outer(border = brd(col, 1.5)) %>% flextable::padding(i = 2, padding.left = 14, padding.top = 8)
  ft <- flextable::height(ft, i = 1, height = hh) %>% flextable::height(i = 2, height = h - hh) %>% flextable::hrule(rule = "exact") %>% flextable::width(width = w)
  ft
}

# rows: data.frame(bar, c1, c2, c3, c4) ; widths in inches
ft_rows <- function(df, colours, widths, h, sizes = c(13, 11, 13, 12), row_bg = c("#EAF1F5", "#F3F6FA")) {
  ft <- flextable::flextable(df) %>% flextable::delete_part("header")
  n <- nrow(df)
  for (i in seq_len(n)) {
    ft <- flextable::compose(ft, i = i, j = 2, value = flextable::as_paragraph(ch(toupper(df[[2]][i]), sizes[1], TRUE, COL$navy)))
    ft <- flextable::compose(ft, i = i, j = 3, value = flextable::as_paragraph(ch(df[[3]][i], sizes[2], FALSE, COL$muted)))
    ft <- flextable::compose(ft, i = i, j = 4, value = flextable::as_paragraph(ch(df[[4]][i], sizes[3], TRUE, COL$text)))
    ft <- flextable::compose(ft, i = i, j = 5, value = flextable::as_paragraph(ch(df[[5]][i], sizes[4], FALSE, COL$text)))
    ft <- flextable::compose(ft, i = i, j = 1, value = flextable::as_paragraph(ch(" ", 8)))
    ft <- flextable::bg(ft, i = i, j = 1, bg = colours[i]) %>% flextable::bg(i = i, j = 2:5, bg = row_bg[(i %% 2) + 1])
  }
  ft <- flextable::border_remove(ft) %>% flextable::set_table_properties(layout = "fixed") %>% flextable::valign(valign = "center") %>%
    flextable::padding(padding.left = 8, padding.right = 8) %>% flextable::border_inner_h(border = brd("white", 6))
  for (j in seq_along(widths)) ft <- flextable::width(ft, j = j, width = widths[j])
  flextable::height(ft, height = h) %>% flextable::hrule(rule = "exact")
}

pp_add <- function(x, obj, l, t, w, h) officer::ph_with(x, obj, location = officer::ph_location(left = l, top = t, width = w, height = h))
pp_txt <- function(x, txt, l, t, w, h, size = 14, bold = FALSE, col = COL$text, align = "left", italic = FALSE) {
  fp <- officer::fpar(officer::ftext(txt, fpt(size, bold, col, italic)), fp_p = officer::fp_par(text.align = align))
  officer::ph_with(x, fp, location = officer::ph_location(left = l, top = t, width = w, height = h))
}
pp_img <- function(x, file, l, t, w, h) {
  if (is.null(file) || !file.exists(file)) return(x)
  officer::ph_with(x, officer::external_img(file, width = w, height = h), location = officer::ph_location(left = l, top = t, width = w, height = h))
}

new_slide <- function(x, title, subtitle, n, header, span) {
  x <- officer::add_slide(x, layout = "Blank", master = "Office Theme")
  x <- pp_txt(x, header, 0.45, 0.12, 8, 0.3, 10.5, FALSE, COL$muted)
  x <- pp_txt(x, span, 9.4, 0.12, 3.5, 0.3, 10.5, FALSE, COL$muted, "right")
  x <- pp_txt(x, title, 0.45, 0.42, 12.4, 0.65, 30, TRUE, COL$navy)
  x <- pp_txt(x, subtitle, 0.45, 1.08, 12.4, 0.4, 15.5, FALSE, COL$muted)
  pp_txt(x, as.character(n), 12.5, 7.1, 0.5, 0.3, 10, FALSE, COL$muted, "right")
}

bg_title_png <- function(file) {
  set.seed(7); n <- 60
  d <- tibble(x = runif(n, 0, 16), y = runif(n, 0, 9), s = runif(n, 2, 14), a = runif(n, 0.05, 0.22))
  g <- expand.grid(x = seq(0, 16, length.out = 240), y = seq(0, 9, length.out = 140)); g$v <- g$x / 16 * 0.5 + g$y / 9 * 0.5
  p <- ggplot() + geom_raster(data = g, aes(x, y, fill = v), interpolate = TRUE) +
    scale_fill_gradient(low = "#0F2A47", high = "#1F5D9A", guide = "none") +
    geom_point(data = d, aes(x, y, size = s, alpha = a), colour = "#7FB2E5", show.legend = FALSE) + scale_size_identity() +
    coord_cartesian(xlim = c(0, 16), ylim = c(0, 9), expand = FALSE) + theme_void()
  ragg::agg_png(file, width = 13.333, height = 7.5, units = "in", res = 110); print(p); dev.off()
}

circles_png <- function(counts, labels, share, cols, file) {
  d <- tibble(y = rev(seq_along(counts)), n = counts, lab = labels, sh = share, col = cols)
  p <- ggplot(d) + geom_point(aes(0.9, y), size = 21, colour = d$col) + geom_text(aes(0.9, y, label = n), colour = "white", fontface = "bold", size = 8) +
    geom_text(aes(1.75, y, label = lab), hjust = 0, size = 6, colour = COL$text) +
    geom_text(aes(5.3, y, label = sh), hjust = 1, size = 4.2, colour = COL$muted) +
    scale_x_continuous(limits = c(0, 5.4), expand = c(0, 0)) + scale_y_continuous(limits = c(0.4, length(counts) + 0.6)) + theme_void() +
    theme(plot.background = element_rect(fill = "#F3F6FA", colour = "#D9E0E8", linewidth = 1))
  ragg::agg_png(file, width = 4.9, height = 4.75, units = "in", res = 200); print(p); dev.off()
}

fmt_lines <- function(df, fmt) lapply(seq_len(nrow(df)), function(i) fmt(df[i, ]))

build_pptx <- function(res, file, chdir) {
  K <- res$K; ct <- res$country; pt <- res$periods; msgs <- res$messages
  span <- sprintf("%s – %s", format(K$from, "%B %Y"), format(K$to, "%B %Y"))
  header <- sprintf("%s | %s", CFG$org, res$title)
  x <- officer::read_pptx(make_template_169())
  n <- 0; ns <- function(title, sub) { n <<- n + 1; new_slide(x, title, sub, n, header, span) }
  per <- pt$periods$period; two <- length(per) >= 2
  scored <- ct %>% filter(!is.na(latest_score))

  # 1 --- title
  tf <- tempfile(fileext = ".png"); bg_title_png(tf)
  x <- officer::add_slide(x, layout = "Blank", master = "Office Theme"); n <- n + 1
  x <- pp_img(x, tf, 0, 0, 13.333, 7.5)
  x <- pp_txt(x, toupper(header), 0.9, 1.7, 11.5, 0.4, 13, TRUE, "#BFD7EE")
  x <- pp_txt(x, toupper(CFG$title), 0.9, 2.2, 11.5, 1.8, 40, TRUE, "white")
  x <- pp_txt(x, "Country scores, trends, data quality, risk prioritisation and decisions", 0.9, 4.15, 11.5, 0.5, 18, FALSE, "#DCE9F7")
  x <- pp_txt(x, sprintf("%s | %d countries | %d campaign rounds scored", span, K$n_countries, K$n_campaigns), 0.9, 6.3, 11.5, 0.4, 15, FALSE, "white")

  # 2 --- method
  x <- ns("Review methodology", "A standardised campaign-round analysis designed to compare countries and guide regional decisions")
  steps <- list(c("1  STRUCTURE", sprintf("Administrative data restricted to %s, grouped by country, campaign and round (district records collapsed).", span)),
                c("2  CLASSIFY", sprintf("%d indicators in 3 pillars: coverage, data quality and LQAS verification.", nrow(IND))),
                c("3  COMPARE", "Each campaign is compared with the standard of every indicator (listed in the workbook)."),
                c("4  CHECK", "Missing, implausible and unmatched values are counted; campaigns that cannot be evaluated stay visible."),
                c("5  ANALYZE", "Trends use the first and the latest evaluable campaign of each country only."),
                c("6  SEGMENT", "Countries are placed in bands and in action segments by rule, never by hand."))
  df <- data.frame(a = map_chr(steps[1:3], 2), b = "", stringsAsFactors = FALSE)
  card_tbl <- function(st) { ft <- flextable::flextable(data.frame(a = "", b = "", c = "")) %>% flextable::delete_part("header")
    for (j in 1:3) ft <- flextable::compose(ft, j = j, value = flextable::as_paragraph(ch(st[[j]][1], 13, TRUE, COL$navy), ch(paste0("\n", st[[j]][2]), 12, FALSE, COL$text)))
    ft <- ft %>% flextable::border_remove() %>% flextable::set_table_properties(layout = "fixed") %>% flextable::bg(bg = "#EEF3F8") %>% flextable::valign(valign = "center") %>%
      flextable::padding(padding.left = 12, padding.right = 12, padding.top = 8) %>% flextable::border_inner_v(border = brd("white", 8))
    ft <- flextable::width(ft, width = 4.1) %>% flextable::height(height = 1.45) %>% flextable::hrule(rule = "exact"); ft }
  x <- pp_add(x, card_tbl(steps[1:3]), 0.45, 1.75, 12.4, 1.45)
  x <- pp_add(x, card_tbl(steps[4:6]), 0.45, 3.35, 12.4, 1.45)
  x <- pp_add(x, ft_banner(sprintf("SCORE = indicators met ÷ indicators evaluable × 100"), 7.4, 0.95, COL$navy, "white", 17), 0.45, 5.15, 7.4, 0.95)
  cau <- flextable::flextable(data.frame(a = "")) %>% flextable::delete_part("header") %>%
    flextable::compose(j = 1, value = flextable::as_paragraph(ch("CAUTION RULES", 11, TRUE, COL$red), ch(sprintf("\nNot-evaluable indicators are left out of the denominator. A campaign with fewer than %d evaluable indicators is N/E and is not scored; rounds of fewer than %d districts are not scored. Scores are rounded to whole numbers.", CFG$min_evaluable, CFG$min_districts), 11, FALSE, COL$text))) %>%
    tbl_base(4.8) %>% flextable::bg(bg = "#FDECEE") %>% flextable::border_outer(border = brd("#F3C6CB")) %>% flextable::valign(valign = "center")
  cau <- flextable::width(cau, width = 4.8) %>% flextable::height(height = 0.95) %>% flextable::hrule(rule = "exact")
  x <- pp_add(x, cau, 8.05, 5.15, 4.8, 0.95)
  x <- pp_txt(x, sprintf("Source: %s%s | Unit of analysis: campaign round | Period: %s", basename(res$admin_file), ifelse(is.null(res$lqas_file), "", paste0(" + ", basename(res$lqas_file))), span), 0.45, 6.35, 12.4, 0.35, 11, FALSE, COL$muted)

  # 3 --- four executive messages (slide-13 layout)
  x <- ns("Four executive messages", "Combined view of performance, operational causes and risks before the next campaigns")
  mb <- K$mean_by; tw <- 2.95; gap <- 0.2
  tiles <- list(list(pct(K$mean_score, 1), COL$navy, "Mean campaign score", sprintf("%d evaluable campaigns", K$n_eval)),
                if (nrow(mb) >= 2) list(pct(mb$mean[2], 1), COL$navy, sprintf("%s mean campaign score", mb$period[2]), sprintf("vs %s in %s", pct(mb$mean[1], 1), mb$period[1])) else list(K$n_countries_scored, COL$navy, "Countries scored", sprintf("of %d", K$n_countries)),
                list(K$n_campaigns, COL$navy, "Campaign rounds", if (nrow(K$ne_by) >= 2) paste(sprintf("%d in %s", K$ne_by$campaigns, K$ne_by$period), collapse = "; ") else ""),
                list(K$n_ne, COL$red, "Campaigns not evaluable", if (nrow(K$ne_by) >= 2) paste(sprintf("%d in %s", K$ne_by$ne, K$ne_by$period), collapse = "; ") else ""))
  for (i in 1:4) x <- pp_add(x, ft_tile(tiles[[i]][[1]], tiles[[i]][[2]], tiles[[i]][[3]], tiles[[i]][[4]], tw, 1.55), 0.45 + (i - 1) * (tw + gap), 1.7, tw, 1.55)
  cc <- c(COL$navy, COL$navy, COL$teal, COL$teal)
  for (i in 1:4) {
    ft <- flextable::flextable(data.frame(a = "", b = "")) %>% flextable::delete_part("header") %>%
      flextable::compose(j = 1, value = flextable::as_paragraph(ch(as.character(i), 18, TRUE, "white"))) %>%
      flextable::compose(j = 2, value = flextable::as_paragraph(ch(msgs[[i]]$title, 15, TRUE, COL$navy), ch(paste0("\n", msgs[[i]]$body), 12, FALSE, COL$text))) %>%
      flextable::border_remove() %>% flextable::set_table_properties(layout = "fixed") %>% flextable::bg(j = 1, bg = cc[i]) %>% flextable::bg(j = 2, bg = ifelse(i %% 2 == 1, "#EAF1F5", "#F3F6FA")) %>%
      flextable::align(j = 1, align = "center") %>% flextable::valign(valign = "center") %>% flextable::padding(padding.left = 10, padding.right = 10)
    ft <- flextable::width(ft, j = 1, width = 0.6) %>% flextable::width(j = 2, width = 5.4) %>% flextable::height(height = 1.55) %>% flextable::hrule(rule = "exact")
    x <- pp_add(x, ft, 0.45 + ((i - 1) %% 2) * 6.2, 3.55 + ((i - 1) %/% 2) * 1.72, 6.0, 1.55)
  }
  x <- pp_txt(x, sprintf("Scores computed from %s by this script (one calculation for every figure). Standards are working defaults listed in the workbook.", basename(res$admin_file)), 0.45, 7.0, 11.8, 0.35, 10.5, FALSE, COL$muted)

  # 4 --- ranking (slide-4 layout)
  x <- ns("Compliance varies widely across countries", sprintf("Latest evaluable campaign: %s to %s | Mean score across %d evaluable campaigns: %s", pct(K$worst$latest_score), pct(K$best$latest_score), K$n_eval, pct(K$mean_score, 1)))
  x <- pp_img(x, file.path(chdir, "01_ranking_latest_score.png"), 0.45, 1.6, 8.6, 5.3)
  x <- pp_add(x, ft_tile(paste0(K$best$latest_score, " %"), COL$green, paste0(K$best$country, " — highest level"), "", 3.4, 1.55), 9.45, 1.7, 3.4, 1.55)
  x <- pp_add(x, ft_tile(paste0(K$worst$latest_score, " %"), COL$red, paste0(K$worst$country, " — lowest level"), "", 3.4, 1.55), 9.45, 3.45, 3.4, 1.55)
  x <- pp_add(x, ft_tile(sprintf("%d / %d", K$n_below, K$n_countries_scored), COL$navy, sprintf("Countries below %d%%", CFG$bands[2]), "", 3.4, 1.55), 9.45, 5.2, 3.4, 1.55)
  x <- pp_txt(x, sprintf("Colours: red <%d%% · orange %d–%d%% · gold %d–%d%% · green ≥%d%%. Latest evaluable campaign of each country, %s.", CFG$bands[1], CFG$bands[1], CFG$bands[2] - 1, CFG$bands[2], CFG$bands[3] - 1, CFG$bands[3], span), 0.45, 6.95, 12, 0.4, 10.5, FALSE, COL$muted)

  # 5 --- trajectories
  imp <- ct %>% filter(trend == "Improving") %>% arrange(desc(change)); dec <- ct %>% filter(trend == "Declining") %>% arrange(change)
  nochg <- ct %>% filter(trend == "No change"); single <- ct %>% filter(trend %in% c("Single evaluable campaign", "No evaluable campaign"))
  x <- ns("Trajectories diverge", ifelse(nrow(imp) && nrow(dec), "Rapid progress exists, but several countries are declining", ifelse(nrow(dec), "Several countries are declining", "Progress is visible in some countries")))
  x <- pp_img(x, file.path(chdir, "02_trajectories.png"), 0.45, 1.6, 8.0, 4.9)
  x <- pp_add(x, ft_panel("IMPROVEMENTS TO CONSOLIDATE", fmt_lines(head(imp, 6), function(r) list(label = paste0(r$country, " "), value = sprintf("%+d", as.integer(r$change)), col = COL$green)), COL$green, 4.3, 2.35), 8.6, 1.7, 4.3, 2.35)
  x <- pp_add(x, ft_panel("DECLINES TO REVERSE", fmt_lines(head(dec, 6), function(r) list(label = paste0(r$country, " "), value = sprintf("%+d", as.integer(r$change)), col = COL$red)), COL$red, 4.3, 2.35), 8.6, 4.2, 4.3, 2.35)
  x <- pp_txt(x, sprintf("No change: %s. Single evaluable campaign (no trend): %s.", list_names(nochg$country, 8), list_names(single$country, 10)), 0.45, 6.55, 8.0, 0.55, 11, FALSE, COL$muted)
  x <- pp_add(x, ft_banner("Each trajectory is the change between the first and the latest evaluable campaign in the period.", 12.4, 0.5, COL$navy, "white", 14), 0.45, 6.95, 12.4, 0.45)

  # 6 --- pillars (slide-7 layout)
  x <- ns("What is driving the compliance score?", sprintf("Share of evaluable campaigns meeting each standard%s", ifelse(two, sprintf(", %s vs %s", per[1], per[2]), sprintf(", %s", per[1]))))
  x <- pp_img(x, file.path(chdir, "03_indicators_by_period.png"), 0.45, 1.6, 7.6, 5.0)
  ind <- pt$ind; lastp <- tail(per, 1)
  wide <- ind %>% select(id, label, period, rate) %>% pivot_wider(names_from = period, values_from = rate)
  if (two) { wide$prev <- wide[[per[1]]]; wide$last <- wide[[per[2]]] } else { wide$prev <- NA_real_; wide$last <- wide[[per[1]]] }
  wide$delta <- wide$last - wide$prev
  improving <- wide %>% filter(!is.na(delta), delta >= CFG$change_flag) %>% arrange(desc(delta)) %>% head(4)
  weak <- wide %>% filter(last < CFG$weak_below, is.na(prev) | prev < CFG$weak_below) %>% arrange(last) %>% head(4)
  corr <- wide %>% filter((!is.na(delta) & delta <= -CFG$change_flag) | last < 25) %>% anti_join(weak %>% filter(last >= 25 & (is.na(delta) | delta > -CFG$change_flag)), by = "id") %>% arrange(last) %>% head(4)
  fl <- function(col) function(r) list(label = paste0(r$label, ": "), value = if (two && !is.na(r$prev)) sprintf("%d%% → %d%%", round(r$prev), round(r$last)) else sprintf("%d%%", round(r$last)), col = col)
  x <- pp_add(x, ft_panel(ifelse(two, sprintf("IMPROVING SIGNALS  (%s → %s)", per[1], per[2]), "IMPROVING SIGNALS"), fmt_lines(improving, fl(COL$green)), COL$green, 4.7, 1.75), 8.2, 1.65, 4.7, 1.75)
  x <- pp_add(x, ft_panel(sprintf("PERSISTENTLY WEAK  (below %d%%%s)", CFG$weak_below, ifelse(two, " in both quarters", "")), fmt_lines(weak, fl(COL$orange)), COL$orange, 4.7, 1.75), 8.2, 3.55, 4.7, 1.75)
  x <- pp_add(x, ft_panel("IMMEDIATE CORRECTIVE ACTION", c(fmt_lines(corr, fl(COL$red)), if (K$n_ne > 0) list(list(label = "Data completeness: ", value = sprintf("%d N/E campaign(s) to complete or justify", K$n_ne), col = COL$red))), COL$red, 4.7, 1.75), 8.2, 5.45, 4.7, 1.55)
  x <- pp_txt(x, sprintf("Standard met = value reaches the threshold of the indicator. %s. N/E results excluded.", paste(sprintf("%s: %d evaluable campaigns", pt$periods$period, (K$mean_by %>% right_join(pt$periods, by = c("period", "period_key")))$n), collapse = "; ")), 0.45, 6.7, 7.6, 0.5, 10.5, FALSE, COL$muted)

  # 6b --- LQAS passed, all campaigns, then the zoom
  lt <- res$lqas_t
  if (length(lt$periods)) {
    a_ <- lt$all; pp <- lt$periods
    x <- ns(sprintf("LQAS: proportion passed, %s", paste(pp, collapse = " vs ")), "Share of LQAS lots passed and of campaigns accepted, all campaigns with LQAS results")
    x <- pp_img(x, file.path(chdir, "06_lqas_passed_all.png"), 0.45, 1.6, 8.6, 5.0)
    big <- function(i) if (i <= nrow(a_) && a_$lots[i] > 0) pct(a_$rate[i]) else "n/a"
    if (all(a_$lots == 0)) {
      x <- pp_add(x, ft_banner(sprintf("No LQAS result in the LQAS file for %s. LQAS data end on %s. Add the recent LQAS file (--lqas) and rerun to fill this slide.", paste(pp, collapse = " and "), ifelse(is.na(K$lqas_max), "n/a", format(K$lqas_max))), 3.4, 3.4, COL$orange, "white", 14), 9.45, 1.7, 3.4, 3.4)
    } else
    x <- pp_add(x, ft_tile(big(nrow(a_)), COL$navy, sprintf("Lots passed in %s", a_$period[nrow(a_)]), sprintf("%s of %s lots", comma(a_$passed[nrow(a_)]), comma(a_$lots[nrow(a_)])), 3.4, 1.55), 9.45, 1.7, 3.4, 1.55)
    if (all(a_$lots == 0)) NULL else if (nrow(a_) >= 2) {
      dlt <- if (all(a_$lots > 0)) a_$rate[2] - a_$rate[1] else NA_real_
      x <- pp_add(x, ft_tile(big(1), COL$muted, sprintf("Lots passed in %s", a_$period[1]), sprintf("%s of %s lots", comma(a_$passed[1]), comma(a_$lots[1])), 3.4, 1.55), 9.45, 3.45, 3.4, 1.55)
      x <- pp_add(x, ft_tile(ifelse(is.na(dlt), "n/a", sprintf("%+.1f pts", dlt)), ifelse(is.na(dlt), COL$muted, ifelse(dlt >= 0, COL$green, COL$red)), "Change in lots passed", sprintf("%s to %s", a_$period[1], a_$period[2]), 3.4, 1.55), 9.45, 5.2, 3.4, 1.55)
    } else x <- pp_add(x, ft_tile(sprintf("%d / %d", a_$accepted[1], a_$campaigns[1]), COL$navy, "Campaigns accepted", sprintf("\u2265%d%% of lots passed", CFG$lqas_pass_std), 3.4, 1.55), 9.45, 3.45, 3.4, 1.55)
    x <- pp_txt(x, sprintf("A lot passes LQAS when it meets the decision rule recorded in the LQAS file. Campaigns are grouped by the quarter of their start date. LQAS data end on %s.", ifelse(is.na(K$lqas_max), "n/a", format(K$lqas_max))), 0.45, 6.75, 12.4, 0.5, 10.5, FALSE, COL$muted)
    zm <- res$zoom %||% CFG$zoom_countries; z_ <- lt$zoom
    x <- ns(sprintf("LQAS zoom: %s", list_names(zm, 6)), sprintf("Proportion of LQAS lots passed, %s", paste(pp, collapse = " vs ")))
    x <- pp_img(x, file.path(chdir, "07_lqas_passed_zoom.png"), 0.45, 1.6, 8.6, 5.0)
    zl <- lapply(zm, function(cn_) {
      r <- z_ %>% filter(country == cn_) %>% arrange(period_key)
      if (!nrow(r)) return(list(label = paste0(cn_, ": "), value = "no LQAS result", col = COL$muted))
      v <- paste(sprintf("%s %d%%", r$period, round(r$rate)), collapse = " \u2192 ")
      list(label = paste0(cn_, ": "), value = v, col = COL$navy)
    })
    x <- pp_add(x, ft_panel("LOTS PASSED BY COUNTRY", zl, COL$navy, 3.8, 3.4), 9.1, 1.7, 3.8, 3.4)
    x <- pp_add(x, ft_banner("Countries with no bar have no LQAS result for that quarter in the LQAS file.", 3.8, 1.3, COL$orange, "white", 13), 9.1, 5.3, 3.8, 1.3)
  }

  # 7 --- bands (slide-8 layout)
  bc <- table(factor(scored$band, levels = levels(scored$band))); bl <- levels(scored$band)
  x <- ns(sprintf("Latest performance %s", ifelse(K$n_below >= 0.5 * K$n_countries_scored, "remains weak in many countries", "is polarised")), sprintf("%d of %d countries have a latest evaluable campaign score below %d%%", K$n_below, K$n_countries_scored, CFG$bands[2]))
  cf <- tempfile(fileext = ".png"); circles_png(as.integer(bc), c("Very high", "High", "Moderate", "Low"), sprintf("%d%% of countries", round(100 * as.integer(bc) / max(1, K$n_countries_scored))), c(COL$red, "#B8651B", "#C9A45C", COL$green), cf)
  x <- pp_img(x, cf, 0.45, 1.65, 4.9, 4.75)
  names_b <- c(paste("Below", CFG$bands[1], "%"), bl[2], bl[3], paste(CFG$bands[3], "% or above"))
  lists_b <- map_chr(1:4, function(.x) { sel <- scored[scored$band == bl[.x], ]; list_names(sel$country[order(-sel$latest_score)], 9) })
  ft <- flextable::flextable(data.frame(a = rep("", 4), b = rep("", 4), c = rep("", 4))) %>% flextable::delete_part("header")
  for (i in 1:4) {
    ft <- flextable::compose(ft, i = i, j = 1, value = flextable::as_paragraph(ch(" ", 8)))
    ft <- flextable::compose(ft, i = i, j = 2, value = flextable::as_paragraph(ch(toupper(names_b[i]), 14, TRUE, COL$navy)))
    ft <- flextable::compose(ft, i = i, j = 3, value = flextable::as_paragraph(ch(lists_b[i], 14, FALSE, COL$text)))
  }
  ft <- ft %>% flextable::border_remove() %>% flextable::set_table_properties(layout = "fixed") %>% flextable::valign(valign = "center") %>%
    flextable::bg(j = 1, bg = c(COL$red, "#B8651B", "#C9A45C", COL$green)) %>% flextable::bg(j = 2:3, bg = "#EAF1F5") %>% flextable::padding(padding.left = 10) %>%
    flextable::border_inner_h(border = brd("white", 7))
  ft <- flextable::width(ft, j = 1, width = 0.22) %>% flextable::width(j = 2, width = 1.6) %>% flextable::width(j = 3, width = 5.4) %>% flextable::height(height = 1.2) %>% flextable::hrule(rule = "exact")
  x <- pp_add(x, ft, 5.6, 1.65, 7.25, 4.8)
  x <- pp_add(x, ft_banner(sprintf("Priority: focus support on countries below %d%%, while consolidating gains and resolving every N/E campaign.", CFG$bands[2]), 12.4, 0.65, COL$navy, "white", 16), 0.45, 6.55, 12.4, 0.65)

  # 8 --- heatmap
  x <- ns("Where each country falls short", "Indicators met (green), not met (red) or not evaluable (grey) in each country's latest evaluable campaign")
  x <- pp_img(x, file.path(chdir, "04_indicator_heatmap.png"), 0.8, 1.55, 11.7, 5.5)

  # 9 --- data quality
  x <- ns("Data quality is an operational risk", "Data gaps may conceal weak performance or create a false impression of improvement")
  nea <- res$camp %>% filter(!small, !evaluable)
  x <- pp_add(x, ft_tile(K$n_campaigns, COL$navy, "Campaign rounds", sprintf("%d evaluable, %d N/E", K$n_eval, K$n_ne), 2.6, 1.4), 0.45, 1.7, 2.6, 1.4)
  x <- pp_add(x, ft_tile(pct(100 * K$n_ne / max(1, K$n_campaigns), 1), COL$red, "N/E share", "campaigns not scored", 2.6, 1.4), 3.25, 1.7, 2.6, 1.4)
  x <- pp_add(x, ft_tile(comma(K$n_missing), COL$orange, "District records incomplete", sprintf("of %s", comma(K$n_districts)), 2.6, 1.4), 0.45, 3.3, 2.6, 1.4)
  x <- pp_add(x, ft_tile(comma(K$n_implaus), COL$orange, sprintf("Coverage above %d%%", CFG$cov_ceiling), "implausible figures", 2.6, 1.4), 3.25, 3.3, 2.6, 1.4)
  x <- pp_img(x, file.path(chdir, "05_data_quality.png"), 6.2, 1.6, 6.7, 3.2)
  ne_txt <- if (nrow(nea)) paste0(nea$country, " (", nea$Round_Add, ", ", nea$month, ")", collapse = "  •  ") else "No campaign is entirely not evaluable."
  x <- pp_txt(x, "Campaigns not evaluable", 0.45, 4.95, 8, 0.35, 14, TRUE, COL$navy)
  x <- pp_txt(x, ne_txt, 0.45, 5.3, 12.4, 0.8, 12.5, FALSE, COL$text)
  x <- pp_txt(x, sprintf("%d campaign code(s) do not match their country; %d evaluable campaign(s) rest on limited evidence.", K$n_code_bad, K$n_limited), 0.45, 6.05, 12.4, 0.4, 12, FALSE, COL$muted)
  x <- pp_add(x, ft_banner(sprintf("Action: complete or justify the data within %s.", CFG$days_validate), 12.4, 0.6, COL$red, "white", 16), 0.45, 6.55, 12.4, 0.6)

  # 10 --- segments (slide-11 layout)
  sg <- res$segments
  x <- ns("A differentiated strategy by segment", "Differentiate reinforced support, close monitoring, consolidation and urgent validation")
  df <- data.frame(a = "", b = sg$segment, c = ifelse(sg$segment == "Urgent validation", paste(sg$n, ifelse(sg$n == 1, "country", "countries")), paste(sg$n, ifelse(sg$n == 1, "country", "countries"))),
                   d = map_chr(sg$countries, ~ list_names(.x, 9)), e = sg$action, stringsAsFactors = FALSE)
  x <- pp_add(x, ft_rows(df, sg$colour, c(0.25, 1.9, 1.15, 5.0, 4.1), 1.0), 0.45, 1.65, 12.4, 5.2)
  both <- intersect(ct$country[ct$segment %in% c("Reinforced support", "Close monitoring", "Targeted follow-up", "Consolidation")], ct$country[ct$needs_validation])
  x <- pp_txt(x, ifelse(length(both), sprintf("%s appear in two segments: they need both their action segment and data validation (urgent validation is cross-cutting).", list_names(both, 8)), "Urgent validation is cross-cutting: a country can be in an action segment and need data validation."), 0.45, 6.9, 12.4, 0.4, 11, FALSE, COL$muted, "left", TRUE)

  # 11 --- priority countries (slide-12 layout)
  pri <- ct %>% filter(segment == "Reinforced support") %>% arrange(latest_score) %>% head(5)
  pri_col <- COL$red; pri_title <- sprintf("%s require reinforced support", ifelse(nrow(pri) == 1, "One country", paste(nrow(pri), "countries")))
  if (!nrow(pri)) {   # nobody is in the strongest-support segment: show the lowest-scoring countries instead
    pri <- ct %>% filter(!is.na(latest_score), latest_score < CFG$bands[2]) %>% arrange(latest_score) %>% head(5)
    pri_col <- COL$orange; pri_title <- if (nrow(pri)) sprintf("%s score below %d%%: priority for close monitoring", ifelse(nrow(pri) == 1, "One country", paste(nrow(pri), "countries")), CFG$bands[2]) else "No country needs reinforced support"
  }
  x <- ns(pri_title, "The response must target the bottlenecks specific to each context")
  if (nrow(pri)) {
    act <- map_chr(seq_len(nrow(pri)), function(i) {
      k <- paste(pri$Country[i], pri$latest_response[i], pri$latest_round[i], sep = "|")
      f <- res$long %>% filter(key == k, result == "No") %>% count(pillar) %>% left_join(count(IND, pillar, name = "tot"), by = "pillar") %>% mutate(share = n / tot) %>% arrange(desc(share), desc(n))
      if (!nrow(f)) "Sustain the current performance" else str_to_sentence(ACTION_BY_PILLAR[[f$pillar[1]]])
    })
    df <- data.frame(a = "", b = pri$country, c = paste0(pri$latest_score, " %", ifelse(pri$latest_is_last, "", "*")), d = act, e = paste("Below standard:", ifelse(nzchar(pri$failed), pri$failed, "none")), stringsAsFactors = FALSE)
    x <- pp_add(x, ft_rows(df, rep(pri_col, nrow(pri)), c(0.25, 1.9, 1.0, 4.6, 4.65), 0.95, sizes = c(14, 15, 12, 11)), 0.45, 1.65, 12.4, 0.95 * nrow(pri))
  } else x <- pp_txt(x, "No country is in the reinforced-support segment and none scores below the second band.", 0.45, 1.8, 12, 0.5, 16, FALSE, COL$muted)
  x <- pp_txt(x, "* The most recent campaign is not evaluable; the score is that of the latest evaluable campaign. Below standard = indicators not met in that campaign.", 0.45, 6.8, 12.4, 0.5, 10.5, FALSE, COL$muted, "left", TRUE)

  # 12 --- decisions table
  x <- ns("Regional decisions before the next campaign round", "Translate findings into decisions, responsibilities, timelines and evidence of closure")
  dd <- res$decisions
  ft <- flextable::flextable(dd) %>% flextable::set_header_labels(Action = "ACTION", Responsible = "RESPONSIBLE", Timeline = "TIMELINE", Evidence = "EVIDENCE / ACCOUNTABILITY") %>%
    flextable::bg(part = "header", bg = COL$navy) %>% flextable::color(part = "header", color = "white") %>% flextable::bold(part = "header") %>% flextable::fontsize(size = 11, part = "header") %>%
    flextable::fontsize(size = 12, part = "body") %>% flextable::font(fontname = FONT, part = "all") %>% flextable::color(color = COL$text, part = "body") %>% flextable::bold(j = 2, part = "body") %>% flextable::align(j = 2:3, align = "center", part = "all") %>%
    flextable::border_remove() %>% flextable::border_inner_h(border = brd("#D9E0E8"), part = "body") %>% flextable::valign(valign = "center", part = "all") %>%
    flextable::padding(padding.left = 8, padding.right = 8, padding.top = 6, padding.bottom = 6) %>% flextable::set_table_properties(layout = "fixed")
  ft <- flextable::width(ft, j = 1, width = 5.6) %>% flextable::width(j = 2, width = 1.6) %>% flextable::width(j = 3, width = 2.2) %>% flextable::width(j = 4, width = 3.0) %>% flextable::height(height = 0.95, part = "body") %>% flextable::hrule(rule = "atleast")
  x <- pp_add(x, ft, 0.45, 1.7, 12.4, 4.6)
  x <- pp_add(x, ft_banner("Accountability rule: every critical gap must have an owner, a deadline and documented evidence of closure.", 12.4, 0.6, COL$navy, "white", 15), 0.45, 6.55, 12.4, 0.6)

  # 13 --- how to read / limits
  x <- ns("How to read this review", "What the score measures, and what it does not")
  lq <- ifelse(is.na(K$lqas_max), "no LQAS data were found", sprintf("LQAS data end on %s; later campaigns have no LQAS result yet and are not penalised", format(K$lqas_max)))
  lim <- list(list(label = "Standards: ", value = "working defaults (workbook, sheet Standards) to be confirmed by the technical team", col = COL$navy),
              list(label = "Coverage: ", value = "administrative coverage, not an independent verification; co-administered vaccines are collapsed by district", col = COL$navy),
              list(label = "LQAS: ", value = lq, col = COL$navy),
              list(label = "Comparisons: ", value = "quarters can include different countries and campaigns; changes are indicative", col = COL$navy),
              list(label = "Not included: ", value = "end-of-campaign independent monitoring (no usable columns in the admin file), financing and microplan timeliness", col = COL$navy))
  x <- pp_add(x, ft_panel("LIMITS TO KEEP IN MIND", lim, COL$navy, 12.4, 3.7, 0.45, 15), 0.45, 1.7, 12.4, 3.6)
  x <- pp_txt(x, sprintf("Files: %s%s. Script: sia_executive_review.R. Generated %s.", basename(res$admin_file), ifelse(is.null(res$lqas_file), "", paste0(" + ", basename(res$lqas_file))), format(Sys.Date())), 0.45, 5.9, 12.4, 0.5, 11.5, FALSE, COL$muted)
  print(x, target = file)
  invisible(file)
}

# ----------------------------------------------------------------------------- 11. MAIN
run_sia_review <- function(admin_file = NULL, lqas_file = NULL, from = NULL, to = NULL, out_dir = NULL, title = NULL, zoom = NULL,
                           make_pptx = TRUE, make_excel = TRUE) {
  t0 <- Sys.time()
  if (!is.null(title)) CFG$title <<- title
  out_dir <- out_dir %||% CFG$out_dir
  if (!dir.exists(out_dir)) {
    ok <- tryCatch(dir.create(out_dir, recursive = TRUE, showWarnings = FALSE), error = function(e) FALSE)
    if (!dir.exists(out_dir)) { out_dir <- file.path(getwd(), "SIA_review_output"); dir.create(out_dir, showWarnings = FALSE, recursive = TRUE) }
  }
  chdir <- file.path(out_dir, "charts"); dir.create(chdir, showWarnings = FALSE)
  dirs <- c(out_dir, file.path(CFG$dashboard_dir, "input"), file.path(getwd(), "input"), getwd())
  admin_path <- find_input(admin_file %||% CFG$admin_file, c("AFRO_admin_data.xlsx", "AFRO_admin_data.rds"), dirs)
  if (is.na(admin_path)) stop("AFRO_admin_data.xlsx / .rds not found. Use --admin <file> or put it in ", out_dir)
  lqas_path <- find_input(lqas_file %||% CFG$lqas_file, c("AFRO_LQAS_combined.rds"), dirs)
  msg("Admin data:", admin_path); msg("LQAS data :", ifelse(is.na(lqas_path), "not found (LQAS indicators will be N/E)", lqas_path))

  adm <- load_admin(admin_path)
  to   <- as_date_safe(to %||% CFG$period_to) %||% min(max(adm$SIA_date, na.rm = TRUE), Sys.Date())
  from <- as_date_safe(from %||% CFG$period_from) %||% (floor_date(to, "quarter") %m-% months(3))
  msg(sprintf("Period %s to %s", from, to))
  prefixes <- prefix_table(adm)
  lqas <- tryCatch(load_lqas(lqas_path), error = function(e) { msg("LQAS file could not be read:", conditionMessage(e)); NULL })
  lqas_max <- if (!is.null(lqas) && nrow(lqas)) suppressWarnings(max(lqas$lqas_date, na.rm = TRUE)) else as.Date(NA)
  if (is.infinite(lqas_max)) lqas_max <- as.Date(NA)

  units <- build_units(adm, from, to)
  if (!nrow(units)) stop("No campaign with a date between ", from, " and ", to, " in the admin file.")
  sc <- score_campaigns(units, lqas, prefixes, lqas_max)
  camp <- sc$camp; long <- sc$long
  ct <- country_table(camp, long); pt <- period_tables(camp, long)
  K <- make_key_numbers(camp, units, ct, pt, from, to, lqas_max, long)
  zoom <- zoom %||% CFG$zoom_countries
  if (length(zoom) == 1 && grepl(",", zoom)) zoom <- trimws(strsplit(zoom, ",")[[1]])
  zoom <- trimws(zoom[nzchar(zoom)]); zoom <- unique(ifelse(toupper(zoom) %in% c("DRC", "CAR"), toupper(zoom), nice_country(toupper(zoom))))
  lt <- lqas_tables(lqas, camp, zoom)
  res <- list(title = CFG$title, admin_file = admin_path, lqas_file = if (is.na(lqas_path)) NULL else lqas_path, K = K, camp = camp, long = long, units = units,
              country = ct, periods = pt, messages = make_messages(K, ct, pt), decisions = make_decisions(ct, K), segments = segment_table(ct), lqas_t = lt, zoom = zoom, lqas_unmatched = NULL)
  msg(sprintf("%d campaign rounds (%d evaluable, %d N/E, %d too small) in %d countries", K$n_campaigns, K$n_eval, K$n_ne, K$n_small, K$n_countries))

  span <- sprintf("%s to %s", format(from, "%b %Y"), format(to, "%b %Y"))
  save_png(ch_ranking(ct, sprintf("Latest evaluable campaign score by country, %s", span)), file.path(chdir, "01_ranking_latest_score.png"), 8.6, max(4.2, 0.27 * K$n_countries_scored + 1.2))
  save_png(ch_trend(ct), file.path(chdir, "02_trajectories.png"), 8, 5)
  save_png(ch_pillar(pt), file.path(chdir, "03_indicators_by_period.png"), 7.6, 5.4)
  save_png(ch_heat(ct, long), file.path(chdir, "04_indicator_heatmap.png"), 11.7, max(4.5, 0.26 * K$n_countries_scored + 1.6))
  save_png(ch_dq(camp, long), file.path(chdir, "05_data_quality.png"), 6.7, 3.2)
  save_png(ch_lqas_all(lt), file.path(chdir, "06_lqas_passed_all.png"), 8.6, 5)
  save_png(ch_lqas_zoom(lt, zoom), file.path(chdir, "07_lqas_passed_zoom.png"), 8.6, 5)
  msg("Charts written to", chdir)

  if (make_excel) { write_workbook(res, file.path(out_dir, "SIA_review_results.xlsx")); msg("Workbook written") }
  write_narrative(res, file.path(out_dir, "SIA_review_narrative.html"), file.path(out_dir, "SIA_review_narrative.md")); msg("Narrative written")
  saveRDS(res, file.path(out_dir, "SIA_review_results.rds"))
  if (make_pptx) {
    ok <- tryCatch({ build_pptx(res, file.path(out_dir, "SIA_review_slides.pptx"), chdir); TRUE }, error = function(e) { msg("PowerPoint failed:", conditionMessage(e)); FALSE })
    if (ok) msg("Slides written")
  }
  msg(sprintf("Done in %.0f s. Output folder: %s", as.numeric(difftime(Sys.time(), t0, units = "secs")), out_dir))
  invisible(res)
}

if (sys.nframe() == 0L) {
  a <- parse_args(commandArgs(trailingOnly = TRUE))
  invisible(run_sia_review(admin_file = a$admin, lqas_file = a$lqas, from = a$from, to = a$to, out_dir = a$out, title = a$title, zoom = a$zoom))
}
