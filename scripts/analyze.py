#!/usr/bin/env python3
"""
Summary statistics for the HCV new-treatment-anticipation / drug-lag analysis.

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


def main():
    ts = pd.read_csv(os.path.join(DATA, "hcv_timeseries.csv")).set_index("fy")
    prod = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries.csv"))
    year_cols = [c for c in prod.columns if c not in ("group", "product")]
    years = [int(c) for c in year_cols]
    y_min, y_max = min(years), max(years)

    daa = prod[prod["group"] == "DAA"][year_cols]
    daa_total = daa.sum(axis=0)
    daa_total.index = years
    daa_peak_year = int(daa_total.idxmax())

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
            "fy_first": float(ts["ribavirin"].loc[y_min]),
            "fy_last": float(ts["ribavirin"].loc[y_max]),
            "first_year_near_zero": int(
                (ts["ribavirin"][ts["ribavirin"] < 0.01 * ts["ribavirin"].loc[y_min]].index.min())
                if (ts["ribavirin"] < 0.01 * ts["ribavirin"].loc[y_min]).any() else -1),
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
        "n_distinct_daa_products": int((daa.sum(axis=1) > 0).sum()),
    }

    with open(os.path.join(RES, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(json.dumps(out, ensure_ascii=False, indent=2))
    return out


if __name__ == "__main__":
    main()
