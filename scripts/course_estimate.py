#!/usr/bin/env python3
"""
Treatment-course sensitivity analysis for the interferon-free DAA group.

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


def main():
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

    total = sum(anchor_courses.values())
    peak_fy = max(anchor_courses, key=anchor_courses.get)
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
        "per_product": per_product,
    }
    with open(os.path.join(RES, "course_estimate.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(json.dumps({k: out[k] for k in out if k != "per_product"},
                     ensure_ascii=False, indent=2))
    return out


if __name__ == "__main__":
    main()
