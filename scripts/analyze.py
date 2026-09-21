#!/usr/bin/env python3
"""
Summary statistics for the HCV interferon-to-DAA transition analysis.

All numbers are computed from data/hcv_timeseries.csv and
data/hcv_product_timeseries.csv (produced by build_dataset.py) and written to
results/summary.json so that the manuscript/report can read them rather than
hard-coding values. No numeric literals describing results live here.

Formal trend models with uncertainty intervals (segmented / exponential
regression with HAC and bootstrap 95% intervals) are computed separately in
its_analysis.py and written to results/its_summary.json.
"""
import json
import os

import pandas as pd

BASE = os.path.join(os.path.dirname(__file__), "..")
DATA = os.path.join(BASE, "data")
RES = os.path.join(BASE, "results")
os.makedirs(RES, exist_ok=True)


def pct_change(series, y0, y1):
    a = float(series.loc[y0])
    b = float(series.loc[y1])
    return None if a == 0 else (b - a) / a * 100.0


def _first_year_zero(series):
    """First fiscal year from which the observed (published) series is zero and
    stays zero. Because NDB omits products outside the top-N of their class and
    shows cells <1,000 as '-', this is the first year the drug fell *below the
    publication threshold*, not evidence of zero dispensing."""
    years = list(series.index)
    for i, y in enumerate(years):
        if all(float(series.loc[yy]) == 0.0 for yy in years[i:]):
            return int(y)
    return -1


def _rebound(series):
    """Largest year-on-year increase after the series first drops below 1% of its
    starting value, so non-monotonic recoveries are reported rather than hidden."""
    years = list(series.index)
    start = float(series.loc[years[0]])
    low = [y for y in years if float(series.loc[y]) < 0.01 * start]
    if not low:
        return None
    best = None
    for y_prev, y_next in zip(years, years[1:]):
        if y_prev < low[0]:
            continue
        a, b = float(series.loc[y_prev]), float(series.loc[y_next])
        if b > a and (best is None or b - a > best["increase"]):
            best = {"from_fy": int(y_prev), "to_fy": int(y_next),
                    "from_value": a, "to_value": b, "increase": b - a,
                    "fold_increase": (b / a) if a > 0 else None}
    return best


def _censoring_block(ts, ts_up, prod, prod_up, year_cols, years):
    """Describe how much of each series could be hidden by NDB publication rules."""
    long = pd.read_csv(os.path.join(DATA, "target_drugs_long.csv"))
    thr = pd.read_csv(os.path.join(DATA, "ndb_publication_thresholds.csv"))
    bounds = pd.read_csv(os.path.join(DATA, "censoring_bounds_long.csv"))
    oral_av = thr[(thr["cls"] == "抗ウイルス剤") & thr["sheet"].str.startswith("内服薬")]
    daa_obs = prod[prod["group"] == "DAA"][year_cols].sum(axis=0).values
    daa_up = prod_up[prod_up["group"] == "DAA"][year_cols].sum(axis=0).values
    pct_hidden = {int(y): (float(u) - float(o)) / float(o) * 100.0 if o > 0 else None
                  for y, o, u in zip(years, daa_obs, daa_up)}

    def _rows(group):
        return {int(fy): {"observed": float(ts[group].loc[fy]),
                          "upper_bound": float(ts_up[group].loc[fy])}
                for fy in years}

    return {
        "rule": ("NDB Open Data lists only the top-N products by dispensed quantity "
                 "within each therapeutic class and care-setting sheet (N = 30, 100 or "
                 "500 depending on edition and class) and shows any cell below 1,000 "
                 "as '-'. Products or cells that are not published are treated as 0 in "
                 "the observed (lower-bound) series; the upper-bound series adds the "
                 "largest quantity each unpublished cell could contain (the smallest "
                 "published total of the class-sheet when the listing cap was reached, "
                 "otherwise 999)."),
        "suppression_threshold": 1000,
        "n_target_rows": int(len(long)),
        "n_suppressed_target_cells": int(long["suppressed"].sum()),
        "n_unpublished_product_sheet_years_bounded": int(len(bounds)),
        "oral_antiviral_listing_cap_by_fy": {
            int(fy): int(g["n_listed"].max()) for fy, g in oral_av.groupby("fy")},
        "oral_antiviral_outpatient_pharmacy_min_listed_total_by_fy": {
            int(r["fy"]): float(r["min_listed_total"])
            for _, r in oral_av[oral_av["sheet"].str.contains("院外")].iterrows()},
        "daa_total_pct_hidden_max_by_fy": pct_hidden,
        "daa_total_pct_hidden_max_overall": max(v for v in pct_hidden.values() if v is not None),
        "series_observed_vs_upper": {g: _rows(g) for g in ts.columns},
    }


def main():
    ts = pd.read_csv(os.path.join(DATA, "hcv_timeseries.csv")).set_index("fy")
    ts_up = pd.read_csv(os.path.join(DATA, "hcv_timeseries_upper.csv")).set_index("fy")
    prod = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries.csv"))
    prod_up = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries_upper.csv"))
    year_cols = [c for c in prod.columns if c not in ("group", "product")]
    years = [int(c) for c in year_cols]
    y_min, y_max = min(years), max(years)

    daa = prod[prod["group"] == "DAA"][year_cols]
    daa_total = daa.sum(axis=0)
    daa_total.index = years
    daa_peak_year = int(daa_total.idxmax())
    daa_total_up = prod_up[prod_up["group"] == "DAA"][year_cols].sum(axis=0)
    daa_total_up.index = years

    out = {
        "data_source": "NDB Open Data editions 1-10 (FY2014-FY2023), 処方薬 性年齢別薬効分類別数量",
        "metric": "総計 (処方数量) national dispensed quantity; units differ across products",
        "fiscal_year_range": [y_min, y_max],
        "peginterferon": {
            "fy_first": float(ts["IFN_peg"].loc[y_min]),
            "fy_last": float(ts["IFN_peg"].loc[y_max]),
            "pct_change_first_to_last": pct_change(ts["IFN_peg"], y_min, y_max),
        },
        "ribavirin": {
            "classification_note": (
                "Reported separately, NOT as part of interferon-based therapy: ribavirin "
                "was co-administered both with peginterferon (interferon-based therapy) "
                "and with sofosbuvir in the interferon-free sofosbuvir + ribavirin "
                "regimen, so its dispensing mixes old and new regimens."),
            "fy_first": float(ts["ribavirin"].loc[y_min]),
            "fy_last": float(ts["ribavirin"].loc[y_max]),
            "first_year_near_zero": int(
                (ts["ribavirin"][ts["ribavirin"] < 0.01 * ts["ribavirin"].loc[y_min]].index.min())
                if (ts["ribavirin"] < 0.01 * ts["ribavirin"].loc[y_min]).any() else -1),
            "first_year_below_publication_threshold": _first_year_zero(ts["ribavirin"]),
            "rebound": _rebound(ts["ribavirin"]),
            "rebound_note": (
                "The apparent year-on-year increase is within the range that unpublished "
                "cells could hide (see censoring.series_observed_vs_upper.ribavirin) and "
                "is therefore not interpretable as a real rebound."),
            "rebound_within_censoring_bound": (
                _rebound(ts["ribavirin"]) is not None and
                float(ts_up["ribavirin"].loc[_rebound(ts["ribavirin"])["from_fy"]])
                >= float(ts["ribavirin"].loc[_rebound(ts["ribavirin"])["to_fy"]])),
            "pct_change_first_to_last": pct_change(ts["ribavirin"], y_min, y_max),
            "upper_bound_fy_last": float(ts_up["ribavirin"].loc[y_max]),
        },
        "interferon_conventional": {
            "role_note": (
                "Not hepatitis-C-specific (also used for other indications); used as a "
                "non-HCV-specific background comparator series."),
            "fy_first": float(ts["IFN_conv"].loc[y_min]),
            "fy_last": float(ts["IFN_conv"].loc[y_max]),
            "pct_change_first_to_last": pct_change(ts["IFN_conv"], y_min, y_max),
        },
        "protease_inhibitors_with_interferon": {
            "role_note": (
                "First-generation NS3/4A protease inhibitors (telaprevir, simeprevir, "
                "vaniprevir) given with peginterferon + ribavirin; markers of "
                "interferon-BASED triple therapy."),
            "fy_first": float(ts["PI_ifn"].loc[y_min]),
            "fy_first_upper_bound": float(ts_up["PI_ifn"].loc[y_min]),
            "first_year_below_publication_threshold": _first_year_zero(ts["PI_ifn"]),
            "upper_bound_fy_last": float(ts_up["PI_ifn"].loc[y_max]),
        },
        "daa_total": {
            "peak_fy": daa_peak_year,
            "peak_value": float(daa_total.loc[daa_peak_year]),
            "fy_first": float(daa_total.loc[y_min]),
            "fy_last": float(daa_total.loc[y_max]),
            "pct_change_peak_to_last": (float(daa_total.loc[y_max]) - float(daa_total.loc[daa_peak_year]))
            / float(daa_total.loc[daa_peak_year]) * 100.0,
            "pct_change_first_to_peak": (float(daa_total.loc[daa_peak_year]) - float(daa_total.loc[y_min]))
            / float(daa_total.loc[y_min]) * 100.0,
        },
        "daa_total_upper_bound": {
            "peak_fy": int(daa_total_up.idxmax()),
            "peak_value": float(daa_total_up.max()),
            "fy_last": float(daa_total_up.loc[y_max]),
            "pct_change_peak_to_last": (float(daa_total_up.loc[y_max]) - float(daa_total_up.max()))
            / float(daa_total_up.max()) * 100.0,
        },
        "n_distinct_daa_products": int((daa.sum(axis=1) > 0).sum()),
        "censoring": _censoring_block(ts, ts_up, prod, prod_up, year_cols, years),
    }

    with open(os.path.join(RES, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(json.dumps(out, ensure_ascii=False, indent=2))
    return out


if __name__ == "__main__":
    main()
