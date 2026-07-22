#!/usr/bin/env python3
"""
Formal trend / segmented-regression analysis with uncertainty intervals for the
HCV new-treatment-anticipation study.

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
  2. IFN-based drugs (peginterferon, conventional IFN, ribavirin): exponential
     decay (log-linear) regression over the fiscal years with positive dispensing,
     giving an annualized decline rate.

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
        "pre_peak_annual_rate_pct": _annual_rate(pre),
        "pre_peak_annual_rate_hac95": [_annual_rate(pre_ci[0]), _annual_rate(pre_ci[1])],
        "pre_peak_annual_rate_boot95": [float(lo[0]), float(hi[0])],
        "post_peak_annual_rate_pct": _annual_rate(post),
        "post_peak_annual_rate_boot95": [float(lo[1]), float(hi[1])],
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
    return {
        "drug": name,
        "fy_used": [int(years.min()), int(years.max())],
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

    out = {
        "design_note": ("NDB begins FY2014, coincident with IFN-free DAA availability; "
                        "no internal pre-intervention baseline, so a conventional pre/post "
                        "ITS is not identifiable. Reported below are descriptive segmented "
                        "and exponential trend models with HAC and bootstrap 95% intervals "
                        "(n=10 annual observations)."),
        "n_annual_observations": len(years),
        "bootstrap_resamples": N_BOOT,
        "daa_segmented": segmented_daa(daa_total),
        "peginterferon_decay": exp_decay(ts["IFN_peg"], "peginterferon"),
        "conventional_ifn_decay": exp_decay(ts["IFN_conv"], "interferon_conventional"),
        "ribavirin_decay": exp_decay(ts["ribavirin"], "ribavirin"),
    }
    with open(os.path.join(RES, "its_summary.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return out


if __name__ == "__main__":
    main()
