#!/usr/bin/env python3
"""
admin_scope_deep_check.py - statistical checks on the admin data, the SCOPE campaign list and their agreement with LQAS. Everything is computed here.

    python py/admin_scope_deep_check.py

A. ADMIN FILE (scope/data/AFRO_admin_data.xlsx, polio rows). Official coverage = TotalNbrVaccPolio / PopPolio (the admin engine adjusts PopPolio; CVPolio is a reference column holding the stored country value; the engine's CV_nOPV2 / CV_bOPV
   use PopPolio, and AFRO-Dash now computes vaccinated / PopPolio itself). Exact rules: stored CVPolio more than 2 points (the portal's tolerance) from vaccinated / PopPolio, with the effect on the dashboard bands, coverage above 150% or above 200%, target population
   far from the country's usual share of the total population, target population or vaccinated counts that jump against the previous campaign of the same district, the same
   (target, vaccinated) pair in 3+ districts of one campaign round (copy-paste), conflicting duplicate rows, impossible dates. Soft: vaccinated equal to the target, round
   numbers (ending in 00), no vaccinated long after the campaign. Rounding: a z-test of the share of counts ending in 00 per country and year. A model (Isolation Forest) marks
   the 1% most unusual rows per country and is tested by planting known corruptions.
B. ADMIN vs SCOPE. Every planned district-round of SCOPE should have an admin row. Completeness is reported for campaigns that started over 90 days ago, with the share of names that could
   be placed on AFROGDB, because a low share makes the completeness unreliable. Admin rows with no planned round, and start dates that differ by more than 7 days, are counted.
C. ADMIN vs LQAS. Rows are joined on district and campaign date (within 7 days). It reports how often LQAS fails where admin coverage is 100% or more, by coverage band and country.
Caveats: only about 27% of Ethiopia's admin names can be placed on AFROGDB (woreda names repeat across zones), so its duplicate, completeness and join figures are not reliable.
Output (reports only, no personal data) in deep_check/: admin_flags_by_country.csv, admin_cv_column_vs_official.csv, admin_record_flags.csv, admin_rounding_by_country_year.csv, admin_scope_completeness.csv,
admin_vs_lqas_by_band.csv, admin_vs_lqas_by_country.csv, admin_vs_lqas_discordant.csv, admin_scope_summary.json
"""
import json, re, unicodedata, warnings, datetime as dt
from pathlib import Path
import numpy as np, pandas as pd
warnings.filterwarnings("ignore")

HERE = Path(__file__).resolve().parent.parent; REPOS = HERE.parent
OUT = HERE / "deep_check"; OUT.mkdir(exist_ok=True)
TODAY = pd.Timestamp(dt.date.today())
ALIAS = {"CAMEROUN": "CAMEROON", "TANZANIA": "UNITED REPUBLIC OF TANZANIA", "REPUBLIQUE CENTRAFRIQUE": "CENTRAL AFRICAN REPUBLIC", "REPUBLIQUE CENTRAFRICAINE": "CENTRAL AFRICAN REPUBLIC"}


def norm(s):
    if s is None or (isinstance(s, float) and np.isnan(s)) or pd.isna(s): return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()


def rnd(x):
    m = re.search(r"(\d+)", str(x)); return m.group(1) if m else ""


def geo_mapper():
    import geopandas as gpd
    cw = pd.read_csv(REPOS / "AFRO-Geo-Reference" / "reference" / "crosswalk_geonames.csv", encoding="utf-8-sig", dtype=str)[["country", "province", "district", "geo_code"]]
    cm = pd.read_csv(REPOS / "AFRO-Geo-Reference" / "reference" / "crosswalk_manual.csv", encoding="utf-8-sig", dtype=str)
    cm = cm[cm.get("decision", pd.Series(dtype=str)).fillna("").ne("reject")] if "decision" in cm else cm
    g = gpd.read_file(REPOS / "AFRO-Geo-Reference" / "data" / "clean" / "africa_districts_simplified.gpkg", ignore_geometry=True)
    g = g.rename(columns={"ADM0_NAME": "country", "ADM1_NAME": "province", "ADM2_NAME": "district", "GEO_CODE": "geo_code"})[["country", "province", "district", "geo_code"]]
    m = pd.concat([cm[["country", "province", "district", "geo_code"]], cw, g]).dropna(subset=["geo_code"])
    m["c"] = m.country.map(norm).replace(ALIAS); m["k3"] = m.c + "|" + m.province.map(norm) + "|" + m.district.map(norm); m["k2"] = m.c + "|" + m.district.map(norm)
    def uniq(col):
        x = m.drop_duplicates([col, "geo_code"]); return x[~x[col].duplicated(keep=False)].set_index(col).geo_code
    u3, u2 = uniq("k3"), uniq("k2")
    def mp(country, province, district):
        c = pd.Series(country).map(norm).replace(ALIAS).values
        k3 = pd.Series(c).astype(str) + "|" + pd.Series(province).map(norm).values + "|" + pd.Series(district).map(norm).values
        k2 = pd.Series(c).astype(str) + "|" + pd.Series(district).map(norm).values
        out = k3.map(u3); return out.fillna(k2.map(u2)).values
    return mp


def main():
    from sklearn.ensemble import IsolationForest
    from scipy.stats import binom, norm as N
    mp = geo_mapper()
    a = pd.read_excel(REPOS / "scope" / "data" / "AFRO_admin_data.xlsx", dtype=object); a.columns = [str(c) for c in a.columns]
    for c in ("PopPolio", "TotalNbrVaccPolio", "CVPolio", "TotPop", "Total_Nbr_0doseVaccPolio_0_59M"): a[c] = pd.to_numeric(a[c], errors="coerce")
    a["date"] = pd.to_datetime(a["SIA_date"], errors="coerce"); a["rep"] = pd.to_datetime(a["Date_Report"], errors="coerce")
    a = a[a.PopPolio > 0].copy().reset_index(drop=True)
    a["geo"] = mp(a.Country, a.Province, a.District); a["rn"] = a.Round_Add.map(rnd)
    a["nk"] = a.Country.map(norm) + "|" + a.Province.map(norm) + "|" + a.District.map(norm)
    n = len(a); pop, vac, cv = a.PopPolio, a.TotalNbrVaccPolio, a.CVPolio
    F = pd.DataFrame(False, index=a.index, columns=["CV_COLUMN_DIFFERS", "COVERAGE_OVER_150", "COVERAGE_OVER_200", "POP_JUMP", "VACC_JUMP", "COPIED_PAIR", "CONFLICTING_DUPLICATE",
                                                     "DATE_IMPLAUSIBLE", "VACC_EQUALS_TARGET", "ROUND_00", "NO_VACC_AFTER_CAMPAIGN", "MODEL_OUTLIER"])
    F["CV_COLUMN_DIFFERS"] = ((cv - vac / pop).abs() > 0.02) & cv.notna() & vac.notna()    # the portal's own tolerance (2 points); official coverage = vaccinated / PopPolio
    F["COVERAGE_OVER_150"] = (vac / pop) > 1.5; F["COVERAGE_OVER_200"] = (vac / pop) > 2
    a = a.assign(_i=np.arange(n)).sort_values(["nk", "date"]); prev = a.groupby("nk")[["PopPolio", "TotalNbrVaccPolio", "geo"]].shift(1)
    pj = a.PopPolio / prev.PopPolio; vj = a.TotalNbrVaccPolio / prev.TotalNbrVaccPolio
    F.loc[a["_i"].values, "POP_JUMP"] = ((pj > 3) | (pj < 1 / 3)).fillna(False).values; F.loc[a["_i"].values, "VACC_JUMP"] = ((vj > 5) | (vj < 1 / 5)).fillna(False).values
    a = a.sort_values("_i").reset_index(drop=True)
    pair = a.Country + "|" + a.Response.astype(str) + "|" + a.Round_Add.astype(str) + "|" + pop.astype(int).astype(str) + "|" + vac.fillna(-1).astype(int).astype(str)
    cnt = pair.map(pair.value_counts()); F["COPIED_PAIR"] = (cnt >= 3) & (pop > 1000)
    dk = a.nk + "|" + a.Admin_3.astype(str) + "|" + a.Vaccine_type.astype(str) + "|" + a.SIA_date.astype(str) + "|" + a.Response.astype(str) + "|" + a.Round_Add.astype(str); dd = vac.groupby(dk).transform("nunique"); F["CONFLICTING_DUPLICATE"] = (dk.map(dk.value_counts()) > 1) & (dd > 1)
    F["DATE_IMPLAUSIBLE"] = (a.date < "2019-01-01") | (a.date > TODAY + pd.Timedelta(days=120)) | ((a.rep.notna()) & ((a.rep < a.date - pd.Timedelta(days=1)) | (a.rep > a.date + pd.Timedelta(days=730))))
    F["VACC_EQUALS_TARGET"] = (vac == pop); F["ROUND_00"] = (vac >= 1000) & (vac % 100 == 0)
    F["NO_VACC_AFTER_CAMPAIGN"] = ((vac.fillna(0) == 0) & (a.date < TODAY - pd.Timedelta(days=90)))
    # rounding by country-year
    rr = []
    for (c, y), g in a.assign(y=a.date.dt.year).groupby(["Country", "y"]):
        v = g.TotalNbrVaccPolio.dropna(); v = v[v >= 1000]
        if len(v) < 100: continue
        k = int((v % 100 == 0).sum()); p0 = 0.01; z = (k - len(v) * p0) / np.sqrt(len(v) * p0 * (1 - p0))
        rr.append(dict(country=c, year=int(y), rows=len(v), ending_in_00=k, expected=round(len(v) * p0, 1), z=round(z, 1), flagged=bool(z > 6 and k >= 10)))
    rr = pd.DataFrame(rr)
    # model with planted check
    medp = pop.groupby(a.nk).transform("median"); medv = vac.groupby(a.nk).transform("median"); cntk = pop.groupby(a.nk).transform("size")
    X = pd.DataFrame({"pop": np.log1p(pop), "vac": np.log1p(vac.fillna(0)), "cv": np.log((vac / pop).clip(0.001, 50).fillna(1)),
                      "pop_dev": np.log(pop / medp).where(cntk >= 3, 0).fillna(0), "vac_dev": np.log((vac.fillna(0) + 1) / (medv.fillna(0) + 1)).where(cntk >= 3, 0).fillna(0),
                      "z0": np.log1p(a.Total_Nbr_0doseVaccPolio_0_59M.fillna(0)), "lag": ((a.rep - a.date).dt.days.clip(-30, 400)).fillna(0)})
    rec_v, rec_p = [], []
    for c, ix in a.groupby("Country").groups.items():
        if len(ix) < 150: continue
        m = IsolationForest(n_estimators=200, max_samples=256, random_state=7).fit(X.loc[ix]); sc = -m.score_samples(X.loc[ix]); thr = np.quantile(sc, 0.99)
        F.loc[ix, "MODEL_OUTLIER"] = sc > thr
        B = X.loc[ix].copy(); B["vac"] += np.log(3); B["vac_dev"] += np.log(3); B["cv"] += np.log(3); rec_v.append(float((-m.score_samples(B) > thr).mean()))          # vaccinated 3 times too high
        B = X.loc[ix].copy(); B["pop"] += np.log(0.3); B["pop_dev"] += np.log(0.3); B["cv"] -= np.log(0.3); rec_p.append(float((-m.score_samples(B) > thr).mean()))       # target 70% too low
    rec_c = rec_v
    F["any_hard"] = F[["COVERAGE_OVER_200", "POP_JUMP", "VACC_JUMP", "COPIED_PAIR", "CONFLICTING_DUPLICATE", "DATE_IMPLAUSIBLE"]].any(axis=1)
    off = vac / pop; band = lambda x: np.select([x >= .95, x >= .80], [">=95%", "80-94%"], "<80%")
    both = cv.notna() & off.notna(); dash_b, off_b = pd.Series(band(cv), index=a.index)[both], pd.Series(band(off), index=a.index)[both]
    imp = pd.DataFrame({"Country": a.Country[both], "differs": F.CV_COLUMN_DIFFERS[both], "band_differs": dash_b != off_b, "cv_ge95": dash_b == ">=95%", "off_ge95": off_b == ">=95%", "cv_ge95_off_lt80": (dash_b == ">=95%") & (off_b == "<80%"),
                        "mean_cv": cv[both], "mean_official": off[both].clip(upper=5)})
    cvo = imp.groupby("Country").agg(rows=("differs", "size"), cv_column_differs_pct=("differs", lambda x: round(100 * x.mean(), 1)), band_differs_pct=("band_differs", lambda x: round(100 * x.mean(), 1)),
                                    dashboard_ge95_pct=("cv_ge95", lambda x: round(100 * x.mean(), 1)), official_ge95_pct=("off_ge95", lambda x: round(100 * x.mean(), 1)), mean_cv_column_pct=("mean_cv", lambda x: round(100 * x.clip(upper=5).mean(), 1)), mean_official_pct=("mean_official", lambda x: round(100 * x.mean(), 1))).reset_index()
    cvo.sort_values("band_differs_pct", ascending=False).to_csv(OUT / "admin_cv_column_vs_official.csv", index=False)
    keys = a[["Country", "Province", "District", "Response", "Round_Add", "SIA_date", "PopPolio", "TotalNbrVaccPolio", "CVPolio", "geo"]]
    keys.join(F)[F.drop(columns="any_hard").any(axis=1)].to_csv(OUT / "admin_record_flags.csv", index=False, encoding="utf-8")
    cols = [c for c in F.columns]
    bc = F.assign(Country=a.Country).groupby("Country").agg(rows=("any_hard", "size"), **{c: (c, "mean") for c in cols}).reset_index()
    bc[cols] = (bc[cols] * 100).round(2); bc.sort_values("any_hard", ascending=False).to_csv(OUT / "admin_flags_by_country.csv", index=False)
    rr.to_csv(OUT / "admin_rounding_by_country_year.csv", index=False)

    # B admin vs scope
    s = pd.read_csv(REPOS / "scope" / "data" / "scope_.csv", low_memory=False); s["geo"] = mp(s.country, s.province, s.district); s["date"] = pd.to_datetime(s.round_start_date, errors="coerce")
    ok = a[a.geo.notna() & a.date.notna()].sort_values("date")[["geo", "date", "Country", "Response", "Round_Add", "PopPolio", "TotalNbrVaccPolio", "CVPolio"]]
    sp = s[s.geo.notna() & s.date.notna()].sort_values("date")[["geo", "date", "country", "response", "roundNumber"]].reset_index(drop=True)
    j = pd.merge_asof(sp, ok.rename(columns={"date": "adate"}).assign(date=lambda d: d.adate).sort_values("date"), on="date", by="geo", tolerance=pd.Timedelta(days=7), direction="nearest")
    sp["reported"] = j["adate"].notna().values
    j30 = pd.merge_asof(sp[["geo", "date"]], ok.rename(columns={"date": "adate"}).assign(date=lambda d: d.adate).sort_values("date"), on="date", by="geo", tolerance=pd.Timedelta(days=30), direction="nearest")
    sp["reported30"] = j30["adate"].notna().values
    old = sp[sp.date < TODAY - pd.Timedelta(days=90)]; old30 = old.reported30.mean()
    comp = old.assign(year=old.date.dt.year).groupby(["country", "year"]).agg(planned=("reported", "size"), reported=("reported", "sum")).reset_index()
    comp["completeness_pct"] = (100 * comp.reported / comp.planned).round(1)
    mapped_a = a.assign(Country_n=a.Country.map(norm).replace(ALIAS)).groupby("Country_n").geo.apply(lambda x: x.notna().mean()).round(3)
    comp["admin_names_placed_pct"] = (100 * comp.country.map(norm).map(mapped_a)).round(1)
    comp["reliable"] = comp.admin_names_placed_pct >= 95
    comp.to_csv(OUT / "admin_scope_completeness.csv", index=False)
    ak = a[a.geo.notna() & a.date.notna()].sort_values("date")
    ja = pd.merge_asof(ak[["geo", "date", "Country"]], sp[["geo", "date"]].assign(sdate=lambda d: d.date).sort_values("date"), on="date", by="geo", tolerance=pd.Timedelta(days=7), direction="nearest")
    not_planned = ja[ja.sdate.isna()]

    # C admin vs lqas
    l = pd.read_parquet(REPOS / "Posit-LQAS-Workflow" / "data" / "final" / "afro_lqas_repositorty.parquet")
    l = l[(l.total_sampled == 60) & l.geo_code.notna()].copy(); l["date"] = pd.to_datetime(l.round_start_date, errors="coerce"); l = l[l.date.notna()].sort_values("date")
    ak2 = ak.assign(admin_cov=ak.TotalNbrVaccPolio / ak.PopPolio)[["geo", "date", "admin_cov", "Country"]]
    jl = pd.merge_asof(l[["geo_code", "date", "country", "province", "district", "status", "total_sampled", "total_vaccinated"]].rename(columns={"geo_code": "geo"}), ak2.rename(columns={"date": "adate"}).assign(date=lambda d: d.adate).sort_values("date"),
                       on="date", by="geo", tolerance=pd.Timedelta(days=7), direction="nearest")
    l["rn"] = l.roundNumber.map(rnd); ak2["rn"] = ak.loc[ak2.index, "Round_Add"].map(rnd).values
    js = pd.merge_asof(l[["geo_code", "rn", "date", "status", "total_sampled", "total_vaccinated"]].rename(columns={"geo_code": "geo"}), ak2.rename(columns={"date": "adate"}).assign(date=lambda d: d.adate).sort_values("date"), on="date", by=["geo", "rn"], tolerance=pd.Timedelta(days=0), direction="nearest")
    js = js[js.admin_cov.notna()]
    from scipy.stats import spearmanr as _sp
    strict = dict(lots=int(len(js)), spearman=round(float(_sp(js.admin_cov.clip(0, 1.5), js.total_sampled - js.total_vaccinated).statistic), 3), fail_pct_admin_100_plus=round(100 * float(js[js.admin_cov >= 1].status.eq("Fail").mean()), 1), fail_pct_admin_below_90=round(100 * float(js[js.admin_cov < .9].status.eq("Fail").mean()), 1))
    jl = jl[jl.admin_cov.notna()].copy(); jl["fail"] = jl.status.eq("Fail")
    bands = [0, .8, .9, .95, 1.0, 1.1, 1.3, 100]; lab = ["<80%", "80-90%", "90-95%", "95-100%", "100-110%", "110-130%", ">130%"]
    jl["band"] = pd.cut(jl["admin_cov"], bands, labels=lab, right=False)
    bb = jl.groupby("band", observed=True).agg(lots=("fail", "size"), lqas_fail_pct=("fail", lambda x: round(100 * x.mean(), 1))).reset_index(); bb.to_csv(OUT / "admin_vs_lqas_by_band.csv", index=False)
    jl["disc"] = (jl["admin_cov"] >= 1.0) & jl.fail
    bcn = jl.groupby("country").agg(lots=("fail", "size"), admin_100_plus=("admin_cov", lambda x: int((x >= 1.0).sum())), of_which_lqas_fail=("disc", "sum")).reset_index()
    bcn["share_fail_among_admin_100_plus"] = (100 * bcn.of_which_lqas_fail / bcn.admin_100_plus.replace(0, np.nan)).round(1); bcn.sort_values("of_which_lqas_fail", ascending=False).to_csv(OUT / "admin_vs_lqas_by_country.csv", index=False)
    jl[jl.disc].assign(admin_coverage_pct=lambda d: (100 * d["admin_cov"]).round(1), lqas_missed=lambda d: d.total_sampled - d.total_vaccinated)[["country", "province", "district", "date", "admin_coverage_pct", "lqas_missed"]] \
        .sort_values("lqas_missed", ascending=False).to_csv(OUT / "admin_vs_lqas_discordant.csv", index=False)
    from scipy.stats import spearmanr
    rho = spearmanr(jl["admin_cov"].clip(0, 1.5), (jl.total_sampled - jl.total_vaccinated)).statistic
    pc = []
    for c, g in jl.groupby("country"):
        if len(g) >= 200: pc.append(dict(country=c, lots=len(g), spearman=round(float(spearmanr(g.admin_cov.clip(0, 1.5), g.total_sampled - g.total_vaccinated).statistic), 3), fail_pct_admin_100_plus=round(100 * float(g[g.admin_cov >= 1].fail.mean()), 1), fail_pct_admin_below_90=round(100 * float(g[g.admin_cov < .9].fail.mean()), 1) if (g.admin_cov < .9).sum() >= 20 else None))
    pd.DataFrame(pc).sort_values("spearman").to_csv(OUT / "admin_vs_lqas_spearman_by_country.csv", index=False)
    summ = dict(coverage_rule="official coverage = TotalNbrVaccPolio / PopPolio (population as adjusted by the admin engine); CVPolio is a reference column; AFRO-Dash now computes the official figure itself", cv_column_vs_official=dict(
                rows=int(len(imp)), differs_more_than_2_points_pct=round(100 * imp.differs.mean(), 1), band_differs_pct=round(100 * imp.band_differs.mean(), 1), dashboard_ge95_pct=round(100 * imp.cv_ge95.mean(), 1),
                official_ge95_pct=round(100 * imp.off_ge95.mean(), 1), rows_ge95_by_cv_but_lt80_official=int(imp.cv_ge95_off_lt80.sum())), strict_join_exact_date_and_round=strict, admin_polio_rows=int(n), admin_names_placed_on_afrogdb_pct=round(100 * a.geo.notna().mean(), 1), hard_flag_pct_excluding_cv_mismatch=round(100 * F.any_hard.mean(), 2),
                flag_pct={c: round(100 * F[c].mean(), 2) for c in cols}, rounding_country_years_flagged=int(rr.flagged.sum()) if len(rr) else 0, rounding_tested=int(len(rr)),
                model_planted_recall_vaccinated_x3=round(float(np.mean(rec_v)), 3) if rec_v else None, model_planted_recall_target_minus_70pct=round(float(np.mean(rec_p)), 3) if rec_p else None,
                scope_completeness_30_day_window_pct=round(100 * float(old30), 1), scope_completeness_reliable_countries_pct=round(100 * comp[comp.reliable].reported.sum() / comp[comp.reliable].planned.sum(), 1), reliable_planned=int(comp[comp.reliable].planned.sum()), scope_planned_old_enough=int(len(old)), scope_reported=int(old.reported.sum()), scope_completeness_pct=round(100 * old.reported.mean(), 1), scope_names_placed_pct=round(100 * s.geo.notna().mean(), 1),
                admin_rows_with_no_planned_round=int(len(not_planned)), admin_rows_placed=int(len(ak)),
                lqas_lots_joined=int(len(jl)), spearman_admin_cov_vs_lqas_missed=round(float(rho), 3), lqas_fail_pct_by_band=bb.set_index("band").lqas_fail_pct.to_dict(),
                admin_100_plus_and_lqas_fail=int(jl.disc.sum()), admin_100_plus=int((jl["admin_cov"] >= 1.0).sum()))
    (OUT / "admin_scope_summary.json").write_text(json.dumps(summ, indent=2, default=str), encoding="utf-8"); print(json.dumps(summ, indent=2, default=str))


if __name__ == "__main__":
    main()
