#!/usr/bin/env python3
"""
Formal trend / segmented-regression analysis with uncertainty intervals for the
HCV interferon-to-DAA transition study.

Design note (important, honest framing)
---------------------------------------
NDB Open Data begins in FY2014, which is essentially the same time the first
interferon(IFN)-free direct-acting-antiviral (DAA) regimen became available in
Japan. There is therefore NO usable pre-intervention baseline inside NDB for the
IFN-free transition, so a conventional pre/post interrupted time series (ITS)
with a clean counterfactual is not identifiable for that event. With only n=10
annual national observations we instead fit descriptive *trend* models and report
uncertainty intervals:

  1. DAA total: a continuous segmented (broken-stick) log-linear regression with a
     knot at the observed peak fiscal year. Estimates the pre-peak and post-peak
     annual multiplicative rates of change and tests whether the slope changes.
     Because the series starts in FY2014 and peaks in FY2015, the pre-knot segment
     rests on only two observations; that count is reported alongside the estimate
     so the pre-peak slope is not read as a trend.
  2. The same segmented model is applied to the combined estimated treatment volume
     (interferon-free DAA courses + peginterferon courses) from course_estimate.py,
     which is the series that speaks to whether the national treated volume rose and
     then fell.
  3. Peginterferon and conventional interferon: exponential decay (log-linear)
     regression over the fiscal years with positive dispensing, giving an annualized
     decline rate. Ribavirin is fitted too but reported separately, because it
     accompanied both interferon-based therapy and the interferon-free sofosbuvir +
     ribavirin regimen and is therefore not a marker of interferon-based therapy.

Uncertainty:
  * Heteroskedasticity- and autocorrelation-consistent (Newey-West / HAC, maxlags=1)
    standard errors from OLS on log dispensed quantity.
  * A residual bootstrap (10,000 resamples) percentile 95% CI as a small-sample
    cross-check, because n=10 makes any single asymptotic CI fragile.

These are descriptive trend models with uncertainty, NOT causal estimates: they
quantify how fast dispensing rose/fell and how uncertain those rates are; they do
not establish that reporting/listing *caused* individual treatment choices.
"""
import json
import os

import numpy as np
import pandas as pd
import statsmodels.api as sm

import course_estimate

BASE = os.path.join(os.path.dirname(__file__), "..")
DATA = os.path.join(BASE, "data")
RES = os.path.join(BASE, "results")
os.makedirs(RES, exist_ok=True)

RNG = np.random.default_rng(20240722)
N_BOOT = 10000


def _annual_rate(slope):
    """Convert a log-scale slope to an annual multiplicative % change."""
    return (np.exp(slope) - 1.0) * 100.0


def _hac_fit(y_log, X, maxlags=1):
    model = sm.OLS(y_log, X)
    return model.fit(cov_type="HAC", cov_kwds={"maxlags": maxlags})


def _boot_ci(y_log, X, fn, n=N_BOOT):
    """Residual bootstrap percentile 95% CI for arbitrary coefficient function fn(res)."""
    base = sm.OLS(y_log, X).fit()
    fitted, resid = np.asarray(base.fittedvalues), np.asarray(base.resid)
    stats = []
    for _ in range(n):
        yb = fitted + RNG.choice(resid, size=len(resid), replace=True)
        rb = sm.OLS(yb, X).fit()
        stats.append(fn(rb))
    stats = np.asarray(stats)
    lo, hi = np.percentile(stats, [2.5, 97.5], axis=0)
    return lo, hi


def segmented_daa(daa_total):
    years = np.asarray(daa_total.index, dtype=float)
    y0 = years.min()
    t = years - y0
    knot = float(daa_total.idxmax()) - y0
    y_log = np.log(daa_total.values.astype(float))
    # continuous broken-stick basis: intercept, t, (t-knot)_+
    seg = np.clip(t - knot, 0.0, None)
    X = np.column_stack([np.ones_like(t), t, seg])
    res = _hac_fit(y_log, X)
    b1, b2 = res.params[1], res.params[2]
    pre, post = b1, b1 + b2
    # HAC CI for pre slope (param 1) and change (param 2)
    ci = res.conf_int()
    pre_ci = (ci[1][0], ci[1][1])
    change_ci = (ci[2][0], ci[2][1])
    # bootstrap CIs on the annual rates
    lo, hi = _boot_ci(y_log, X, lambda r: [_annual_rate(r.params[1]),
                                           _annual_rate(r.params[1] + r.params[2])])
    return {
        "knot_fy": int(daa_total.idxmax()),
        "n_obs_up_to_knot": int((t <= knot).sum()),
        "pre_peak_caution": (
            "The pre-peak segment is informed by only the observations up to the "
            "knot; with a knot in the second fiscal year this is a two-point "
            "increment, not an estimated trend, and its interval is correspondingly "
            "wide. Reported for completeness, not for interpretation as a rate."),
        "pre_peak_annual_rate_pct": _annual_rate(pre),
        "pre_peak_annual_rate_hac95": [_annual_rate(pre_ci[0]), _annual_rate(pre_ci[1])],
        "pre_peak_annual_rate_boot95": [float(lo[0]), float(hi[0])],
        "post_peak_annual_rate_pct": _annual_rate(post),
        "post_peak_annual_rate_boot95": [float(lo[1]), float(hi[1])],
        "n_obs_from_knot": int((t >= knot).sum()),
        "slope_change_logunits": float(b2),
        "slope_change_hac95": [float(change_ci[0]), float(change_ci[1])],
        "slope_change_p": float(res.pvalues[2]),
        "n_obs": int(len(t)),
    }


def exp_decay(series, name):
    s = series[series > 0].astype(float)
    years = np.asarray(s.index, dtype=float)
    t = years - years.min()
    y_log = np.log(s.values)
    X = np.column_stack([np.ones_like(t), t])
    res = _hac_fit(y_log, X)
    slope = res.params[1]
    ci = res.conf_int()[1]
    lo, hi = _boot_ci(y_log, X, lambda r: _annual_rate(r.params[1]))
    increases = [int(b) for a_, b in zip(s.index, s.index[1:])
                 if float(s.loc[b]) > float(s.loc[a_])]
    ends_early = years.max() < float(np.asarray(series.index, dtype=float).max())
    return {
        "drug": name,
        "fy_used": [int(years.min()), int(years.max())],
        "fy_used_note": (
            "Fiscal years with published (non-suppressed) dispensing; from the "
            "following fiscal year onward the drug was below the NDB publication "
            "threshold (not listed or shown as '-')." if ends_early else
            "Fiscal years with positive dispensing; the series remains positive "
            "through the final fiscal year."),
        "years_with_year_on_year_increase": increases,
        "annual_change_pct": _annual_rate(slope),
        "annual_change_hac95": [_annual_rate(ci[0]), _annual_rate(ci[1])],
        "annual_change_boot95": [float(lo), float(hi)],
        "annual_change_p": float(res.pvalues[1]),
        "n_obs": int(len(t)),
    }


def main():
    ts = pd.read_csv(os.path.join(DATA, "hcv_timeseries.csv")).set_index("fy")
    prod = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries.csv"))
    year_cols = [c for c in prod.columns if c not in ("group", "product")]
    years = [int(c) for c in year_cols]
    daa_total = prod[prod["group"] == "DAA"][year_cols].sum(axis=0)
    daa_total.index = years
    # Recomputed here rather than read from results/course_estimate.json so that
    # the combined-volume model can never pick up a stale file from an earlier run.
    ce = course_estimate.main(verbose=False)
    cv = ce["combined_treatment_volume"]["estimated_courses_by_fy"]
    combined = pd.Series({int(k): float(v) for k, v in cv.items()}).sort_index()

    # Upper-bound sensitivity: unpublished NDB cells filled at their maximum.
    prod_up = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries_upper.csv"))
    ts_up = pd.read_csv(os.path.join(DATA, "hcv_timeseries_upper.csv")).set_index("fy")
    daa_total_up = prod_up[prod_up["group"] == "DAA"][year_cols].sum(axis=0)
    daa_total_up.index = years
    cv_up = ce["censoring_upper_bound"]["combined_estimated_courses_by_fy"]
    combined_up = pd.Series({int(k): float(v) for k, v in cv_up.items()}).sort_index()

    out = {
        "design_note": ("NDB begins FY2014, coincident with IFN-free DAA availability; "
                        "no internal pre-intervention baseline, so a conventional pre/post "
                        "ITS is not identifiable. Reported below are descriptive segmented "
                        "and exponential trend models with HAC and bootstrap 95% intervals "
                        "(n=10 annual observations)."),
        "n_annual_observations": len(years),
        "bootstrap_resamples": N_BOOT,
        "daa_segmented": segmented_daa(daa_total),
        "combined_volume_segmented": segmented_daa(combined),
        "combined_volume_note": (
            "Segmented model on the combined estimated treatment volume "
            "(interferon-free DAA courses + peginterferon courses) from "
            "results/course_estimate.json. Estimates, not observed patient counts."),
        "peginterferon_decay": exp_decay(ts["IFN_peg"], "peginterferon"),
        "conventional_ifn_decay": exp_decay(ts["IFN_conv"], "interferon_conventional"),
        "ribavirin_decay": exp_decay(ts["ribavirin"], "ribavirin"),
        "censoring_upper_bound": {
            "note": ("Same models refitted to the upper-bound series in which every "
                     "NDB cell that is unpublished (product outside the top-N of its "
                     "class, or total shown as '-') is set to the largest value it could "
                     "hide. Gives the least steep post-peak decline compatible with the "
                     "publication rules; a bound, not an estimate."),
            "daa_segmented": segmented_daa(daa_total_up),
            "combined_volume_segmented": segmented_daa(combined_up),
            "peginterferon_decay": exp_decay(ts_up["IFN_peg"], "peginterferon"),
        },
        "ribavirin_note": (
            "Ribavirin accompanied both peginterferon-based therapy and the "
            "interferon-free sofosbuvir + ribavirin regimen, so this decay rate "
            "describes the combined decline of both regimens and is not evidence "
            "about interferon-based therapy alone."),
    }
    with open(os.path.join(RES, "its_summary.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return out


if __name__ == "__main__":
    main()
