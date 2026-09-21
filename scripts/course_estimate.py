#!/usr/bin/env python3
"""
Treatment-course sensitivity analysis for interferon-free DAAs and for
peginterferon-based therapy, plus the combined national treatment volume.

NDB reports national dispensed quantity (tablets/capsules), NOT patient counts.
To give a rough sense of the *practical patient scale*, this script converts the
dispensed quantity of each interferon-free DAA product to an approximate number
of full treatment courses, using explicit, documented per-course unit counts
(daily dose x standard duration) from data/daa_course_assumptions.csv (based on
the Japanese package inserts and the JSH hepatitis C treatment guideline).

Two-drug regimens (daclatasvir+asunaprevir, elbasvir+grazoprevir) are counted
once per regimen using a single "anchor" component so that co-administered drugs
are not double-counted. A duration-sensitivity range is also produced for
glecaprevir/pibrentasvir (8 vs 12 weeks) and sofosbuvir/velpatasvir (12 vs 24
weeks).

Peginterferon dispensing (prefilled syringes, one per weekly dose) is converted
the same way using data/ifn_course_assumptions.csv (48-week standard regimen,
with 24 weeks as the sensitivity alternative). Because the first-generation
NS3/4A protease inhibitors were always given on top of peginterferon, counting
peginterferon courses already covers those patients and avoids double-counting.
Ribavirin is excluded from course counting altogether: it accompanied both
peginterferon-based therapy and the interferon-free sofosbuvir + ribavirin
regimen, so it cannot be attributed to one side.

Putting the two groups on the common scale of estimated courses lets the combined
national hepatitis-C antiviral treatment volume, and the DAA share of it, be
plotted on one axis, which tablets-versus-syringes units do not permit.

ALL outputs are explicitly labelled ESTIMATES and are written to
results/course_estimate.json. They are not observed patient counts.
"""
import json
import os

import pandas as pd

BASE = os.path.join(os.path.dirname(__file__), "..")
DATA = os.path.join(BASE, "data")
RES = os.path.join(BASE, "results")
os.makedirs(RES, exist_ok=True)

# alternative (longer) durations for a sensitivity range, in weeks
ALT_WEEKS = {
    "glecaprevir/pibrentasvir": 12,   # 8w baseline (naive) vs 12w
    "sofosbuvir/velpatasvir": 24,     # 12w baseline vs 24w (prior DAA failure)
}


def _scenario(prod, ts, a, ifn_asmp, years, year_cols):
    """Course estimates for one set of dispensed-quantity series."""
    daa = prod[prod["group"] == "DAA"].set_index("product")
    anchor = {y: 0.0 for y in years}
    for p, row in daa.iterrows():
        if p in a.index and int(a.loc[p, "is_anchor"]) == 1:
            upc = float(a.loc[p, "units_per_course"])
            for c in year_cols:
                anchor[int(c)] += float(row[c]) / upc
    upc_ifn = float(ifn_asmp.loc["peginterferon", "units_per_course"])
    ifn = {y: float(ts["IFN_peg"].loc[y]) / upc_ifn for y in years}
    combined = {y: anchor[y] + ifn[y] for y in years}
    share = {y: (anchor[y] / combined[y] if combined[y] > 0 else None) for y in years}
    peak = max(combined, key=combined.get)
    return {
        "daa_estimated_courses_by_fy": {int(y): round(anchor[y], 1) for y in years},
        "combined_estimated_courses_by_fy": {int(y): round(combined[y], 1) for y in years},
        "combined_peak_fy": int(peak),
        "combined_peak_estimated_courses": round(combined[peak], 1),
        "combined_pct_change_peak_to_last": round(
            (combined[years[-1]] - combined[peak]) / combined[peak] * 100.0, 1),
        "daa_share_by_fy": {int(y): (round(share[y], 4) if share[y] is not None else None)
                            for y in years},
    }


def main(verbose=True):
    prod = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries.csv"))
    asmp = pd.read_csv(os.path.join(DATA, "daa_course_assumptions.csv"))
    year_cols = [c for c in prod.columns if c not in ("group", "product")]
    years = [int(c) for c in year_cols]

    a = asmp.set_index("product")
    daa = prod[prod["group"] == "DAA"].set_index("product")

    per_product = {}
    anchor_courses = {y: 0.0 for y in years}
    anchor_courses_hi = {y: 0.0 for y in years}
    for p, row in daa.iterrows():
        if p not in a.index:
            continue
        upc = float(a.loc[p, "units_per_course"])
        is_anchor = int(a.loc[p, "is_anchor"]) == 1
        upc_alt = upc
        if p in ALT_WEEKS:
            upc_alt = float(a.loc[p, "tablets_per_day"]) * ALT_WEEKS[p] * 7.0
        per_product[p] = {
            "units_per_course": upc,
            "is_anchor": is_anchor,
            "courses_by_fy": {int(c): float(row[c]) / upc for c in year_cols},
        }
        if is_anchor:
            for c in year_cols:
                anchor_courses[int(c)] += float(row[c]) / upc
                anchor_courses_hi[int(c)] += float(row[c]) / upc_alt

    # peginterferon-based therapy on the same estimated-course scale
    ifn_asmp = pd.read_csv(os.path.join(DATA, "ifn_course_assumptions.csv")).set_index("product")
    ts = pd.read_csv(os.path.join(DATA, "hcv_timeseries.csv")).set_index("fy")
    upc_ifn = float(ifn_asmp.loc["peginterferon", "units_per_course"])
    upc_ifn_alt = float(ifn_asmp.loc["peginterferon", "units_per_course_alt"])
    ifn_courses = {y: float(ts["IFN_peg"].loc[y]) / upc_ifn for y in years}
    ifn_courses_alt = {y: float(ts["IFN_peg"].loc[y]) / upc_ifn_alt for y in years}

    # combined national volume and DAA share, both on the estimated-course scale
    combined = {y: anchor_courses[y] + ifn_courses[y] for y in years}
    daa_share = {y: (anchor_courses[y] / combined[y] if combined[y] > 0 else None)
                 for y in years}

    total = sum(anchor_courses.values())
    peak_fy = max(anchor_courses, key=anchor_courses.get)
    combined_peak_fy = max(combined, key=combined.get)

    # Upper-bound scenario: every unpublished NDB cell (product outside the top-N of
    # its class, or a total shown as '-') filled with the largest value it could hide.
    prod_up = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries_upper.csv"))
    ts_up = pd.read_csv(os.path.join(DATA, "hcv_timeseries_upper.csv")).set_index("fy")
    upper = _scenario(prod_up, ts_up, a, ifn_asmp, years, year_cols)
    upper["note"] = (
        "ESTIMATE and BOUND only. Same course conversion applied to the upper-bound "
        "dispensing series from build_dataset.py, in which every cell that NDB Open "
        "Data does not publish (product below the class listing cap, or total shown "
        "as '-') is set to the maximum quantity it could contain. Actual values lie "
        "between the baseline and this bound.")
    upper["daa_share_difference_vs_baseline_pp"] = {
        int(y): (round((upper["daa_share_by_fy"][int(y)] - daa_share[y]) * 100.0, 2)
                 if daa_share[y] is not None and upper["daa_share_by_fy"][int(y)] is not None
                 else None) for y in years}
    out = {
        "note": ("ESTIMATE only. Approximate interferon-free DAA treatment "
                 "courses = dispensed quantity / units_per_course, summed over "
                 "one anchor product per regimen to avoid double-counting "
                 "co-administered drugs. Not observed patient counts."),
        "assumptions_source": ("data/daa_course_assumptions.csv (Japanese "
                               "package inserts; JSH hepatitis C guideline)"),
        "estimated_courses_by_fy": {int(y): round(anchor_courses[y], 1) for y in years},
        "estimated_courses_by_fy_longer_duration": {
            int(y): round(anchor_courses_hi[y], 1) for y in years},
        "estimated_total_courses_fy2014_2023": round(total, 1),
        "estimated_total_courses_fy2014_2023_range": [
            round(sum(anchor_courses_hi.values()), 1), round(total, 1)],
        "peak_fy": int(peak_fy),
        "peak_estimated_courses": round(anchor_courses[peak_fy], 1),
        "peginterferon": {
            "note": ("ESTIMATE only. Peginterferon courses = dispensed syringes / "
                     "units_per_course (one syringe per weekly dose). Includes the "
                     "patients who also received a first-generation protease "
                     "inhibitor, since those were always added to peginterferon."),
            "assumptions_source": "data/ifn_course_assumptions.csv",
            "units_per_course": upc_ifn,
            "units_per_course_alt": upc_ifn_alt,
            "estimated_courses_by_fy": {int(y): round(ifn_courses[y], 1) for y in years},
            "estimated_courses_by_fy_shorter_regimen": {
                int(y): round(ifn_courses_alt[y], 1) for y in years},
            "estimated_total_courses": round(sum(ifn_courses.values()), 1),
            "estimated_total_courses_range": [
                round(sum(ifn_courses.values()), 1),
                round(sum(ifn_courses_alt.values()), 1)],
        },
        "combined_treatment_volume": {
            "note": ("ESTIMATE only. Interferon-free DAA courses + peginterferon "
                     "courses, i.e. an approximate national hepatitis-C antiviral "
                     "treatment volume on a common scale. Ribavirin is excluded "
                     "because it accompanied both interferon-based and "
                     "interferon-free regimens."),
            "estimated_courses_by_fy": {int(y): round(combined[y], 1) for y in years},
            "peak_fy": int(combined_peak_fy),
            "peak_estimated_courses": round(combined[combined_peak_fy], 1),
            "fy_first_estimated_courses": round(combined[years[0]], 1),
            "fy_last_estimated_courses": round(combined[years[-1]], 1),
            "pct_change_peak_to_last": round(
                (combined[years[-1]] - combined[combined_peak_fy])
                / combined[combined_peak_fy] * 100.0, 1),
            "daa_share_by_fy": {int(y): (round(daa_share[y], 4)
                                         if daa_share[y] is not None else None)
                                for y in years},
        },
        "censoring_upper_bound": upper,
        "per_product": per_product,
    }
    with open(os.path.join(RES, "course_estimate.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    if verbose:
        print(json.dumps({k: out[k] for k in out if k != "per_product"},
                         ensure_ascii=False, indent=2))
    return out


if __name__ == "__main__":
    main()
