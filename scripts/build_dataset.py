#!/usr/bin/env python3
"""
Build the HCV drug-lag time series from NDB Open Data (editions 1-10, FY2014-FY2023).

Reads the raw NDB "処方薬 性年齢別薬効分類別数量" workbooks (内服/外用/注射,
外来院内・外来院外・入院) downloaded under data/ndb_raw/dai{N}/f00..f04.xlsx and
extracts the national total dispensed quantity (総計 / 処方数量) for the
hepatitis-C direct-acting antivirals (DAAs) and the interferon-based backbone
(peginterferon, conventional interferon) plus ribavirin.

Ribavirin is deliberately kept in its own group and is NOT counted as part of the
interferon-based backbone: in Japan it was co-administered both with peginterferon
(interferon-based dual/triple therapy) and with sofosbuvir in the interferon-free
sofosbuvir + ribavirin regimen for genotype 2. Its dispensing therefore tracks a
mixture of the old and the new regimens and is not a clean marker of
interferon-based therapy.

Output:
  data/target_drugs_long.csv     one row per (edition, product, formulation, setting)
  data/hcv_timeseries.csv        group x fiscal-year national totals
  data/hcv_product_timeseries.csv product x fiscal-year national totals
  data/ndb_publication_thresholds.csv  per (fy, sheet, therapeutic class): number of
                                 products listed, the edition's listing cap, and the
                                 smallest listed total (the censoring threshold)
  data/censoring_bounds_long.csv one row per (product, fy, sheet) cell that is not
                                 published, with the maximum quantity it could hide
  data/hcv_product_timeseries_upper.csv / data/hcv_timeseries_upper.csv
                                 observed totals plus those maxima (upper-bound series)

Publication rules of NDB Open Data (stated in each workbook header and verified
empirically from the number of products per class):
  * each sheet (oral / topical / injectable x inpatient / outpatient in-hospital /
    outpatient pharmacy) lists only the top-N products by dispensed quantity within
    each therapeutic class (N = 30, 100 or 500 depending on edition and class, with
    ties occasionally pushing the count a few above N); a product below that rank
    is simply absent;
  * any cell (including the national total) below 1,000 is shown as "-".
Absence or "-" therefore means "below the publication threshold", not zero. The
observed series treat such cells as 0 (a lower bound); the *_upper.csv files add
the largest quantity each unpublished cell could contain (the smallest listed
total in that sheet-class when the cap was reached, otherwise 999), for years from
the product's first listing onward.

Metric note: NDB Open Data reports 処方数量 (dispensed quantity: tablets/capsules
for oral drugs, pre-filled syringes/vials for injections), NOT patient counts.
Quantities are therefore comparable within a product/formulation over time but are
not additive across products with different dosage units. Estimated treatment
courses (an approximate patient-count proxy) are derived separately in analyze.py
using documented regimen durations, and are clearly labelled as estimates.
"""
import glob
import os

import openpyxl
import pandas as pd

ROOT = os.path.join(os.path.dirname(__file__), "..", "data", "ndb_raw")
OUT = os.path.join(os.path.dirname(__file__), "..", "data")

EDITION_FY = {1: 2014, 2: 2015, 3: 2016, 4: 2017, 5: 2018,
              6: 2019, 7: 2020, 8: 2021, 9: 2022, 10: 2023}

# Cells below this value are shown as "-" in every NDB Open Data edition.
SUPPRESSION_THRESHOLD = 1000.0
# Listing caps used by NDB Open Data (products per class per sheet); ties may add a
# few rows beyond the cap.
LISTING_CAPS = (30, 100, 500)
CAP_TIE_TOLERANCE = 3


def listing_cap(n_listed):
    """Return the cap a class reached (n within [cap, cap+tolerance]) or None."""
    for c in LISTING_CAPS:
        if c <= n_listed <= c + CAP_TIE_TOLERANCE:
            return c
    return None


# Formulation sheets in which each drug group can appear.
GROUP_SHEET_PREFIX = {"DAA": "内服薬", "PI_ifn": "内服薬", "ribavirin": "内服薬",
                      "IFN_peg": "注射薬", "IFN_conv": "注射薬"}

# Interferon-free DAA products (brand-name substrings as they appear in NDB) -> INN label
DAA = {
    "ソバルディ": "sofosbuvir",
    "ハーボニー": "ledipasvir/sofosbuvir",
    "ダクルインザ": "daclatasvir",
    "スンベプラ": "asunaprevir",
    "マヴィレット": "glecaprevir/pibrentasvir",
    "エプクルーサ": "sofosbuvir/velpatasvir",
    "エレルサ": "elbasvir",
    "グラジナ": "grazoprevir",
    "ヴィキラックス": "ombitasvir/paritaprevir/ritonavir",
}
# First-generation NS3/4A protease inhibitors used *with* peginterferon+ribavirin
# (interferon-BASED triple therapy, not interferon-free). Kept separate so the
# DAA group cleanly represents interferon-free regimens.
PI_IFN = {
    "テラビック": "telaprevir",
    "ソブリアード": "simeprevir",
    "バニヘップ": "vaniprevir",
}
IFN_PEG = ["ペガシス", "ペグイントロン"]
IFN_CONV = ["スミフェロン", "フエロン", "イントロンＡ", "オーアイエフ"]
# Ribavirin: companion drug of BOTH peginterferon-based therapy and the
# interferon-free sofosbuvir + ribavirin regimen. Reported as its own group.
RIBAVIRIN = ["レベトール", "コペガス"]


def classify(name: str):
    for k, v in DAA.items():
        if k in name:
            return "DAA", v
    for k, v in PI_IFN.items():
        if k in name:
            return "PI_ifn", v
    if any(k in name for k in IFN_PEG):
        return "IFN_peg", "peginterferon"
    if any(k in name for k in IFN_CONV):
        return "IFN_conv", "interferon_conventional"
    if any(k in name for k in RIBAVIRIN):
        return "ribavirin", "ribavirin"
    return None, None


def header_cols(ws):
    """Locate header row and column indices; NDB layout varies across editions."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True)):
        vals = [str(c).replace("\n", "") if c is not None else "" for c in row]
        if "薬効分類名称" in vals and any(v.startswith("総計") for v in vals):
            idx = {}
            for j, v in enumerate(vals):
                if v == "薬効分類名称":
                    idx["cls"] = j
                elif v == "医薬品名":
                    idx["name"] = j
                elif v == "単位":
                    idx["unit"] = j
                elif v.startswith("総計"):
                    idx["tot"] = j
            return i + 1, idx
    return None, None


def _to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def publication_bounds(df, thresholds):
    """Maximum quantity hidden in each unpublished (product, fy, sheet) cell.

    For every target product and every fiscal year from its first listing onward,
    each formulation sheet (inpatient / outpatient in-hospital / outpatient pharmacy)
    either shows a numeric total, shows "-" (total < 1,000), or omits the product.
    An omitted product can hide at most the smallest listed total of its class in
    that sheet when the class hit the edition's listing cap; if the cap was not
    reached, omission means every cell was below 1,000.
    """
    thr = thresholds.set_index(["fy", "sheet", "cls"]).sort_index()
    present = df.set_index(["fy", "product", "sheet"]).sort_index()
    fys = sorted(df["fy"].unique())
    rows = []
    for (grp, prod), sub in df.groupby(["group", "product"]):
        cls = sub["yakko_class"].mode().iloc[0]
        first_fy = int(sub["fy"].min())
        prefix = GROUP_SHEET_PREFIX[grp]
        for fy in fys:
            if fy < first_fy:
                continue
            sheets = [s for (f, s, c) in thr.index if f == fy and c == cls
                      and s.startswith(prefix)]
            for sh in sheets:
                key = (fy, prod, sh)
                if key in present.index:
                    cell = present.loc[[key]]
                    if not bool(cell["suppressed"].any()):
                        continue
                    bound, reason = SUPPRESSION_THRESHOLD - 1, "total shown as '-' (<1,000)"
                else:
                    t = thr.loc[(fy, sh, cls)]
                    if bool(t["cap_hit"]):
                        mn = t["min_listed_total"]
                        bound = float(mn) if pd.notna(mn) else SUPPRESSION_THRESHOLD - 1
                        reason = f"not among top {int(t['cap'])} listed products"
                    else:
                        bound, reason = SUPPRESSION_THRESHOLD - 1, "not listed; class below cap (<1,000)"
                rows.append([fy, grp, prod, sh, cls, bound, reason])
    return pd.DataFrame(rows, columns=["fy", "group", "product", "sheet", "yakko_class",
                                       "max_hidden_qty", "reason"])


def main():
    rows = []
    listing = []
    for ed in range(1, 11):
        # Read every workbook for the edition (not just f00-f04): additional
        # medical/dental supplement files are harmless (they contain none of the
        # target drugs) but scanning them all avoids silently dropping any file
        # if the edition's file layout changes.
        for fn in sorted(glob.glob(os.path.join(ROOT, f"dai{ed}", "f*.xlsx"))):
            wb = openpyxl.load_workbook(fn, read_only=True)
            for sh in wb.sheetnames:
                ws = wb[sh]
                hrow, idx = header_cols(ws)
                if not idx:
                    raise RuntimeError(f"header not found: dai{ed} {sh}")
                cur_cls = None
                for r in ws.iter_rows(min_row=hrow + 1, values_only=True):
                    cls = r[idx["cls"]] if idx["cls"] < len(r) else None
                    if cls:
                        cur_cls = cls
                    nm = r[idx["name"]] if idx["name"] < len(r) else None
                    if not nm:
                        continue
                    tot_raw = r[idx["tot"]] if idx["tot"] < len(r) else None
                    tot = _to_float(tot_raw)
                    listing.append([EDITION_FY[ed], sh, str(cur_cls), tot])
                    grp, prod = classify(str(nm))
                    if grp is None:
                        continue
                    unit = r[idx["unit"]] if "unit" in idx and idx["unit"] < len(r) else ""
                    suppressed = tot is None
                    rows.append([ed, EDITION_FY[ed], grp, prod, str(nm),
                                 str(cur_cls), unit, sh, 0.0 if suppressed else tot,
                                 suppressed])
            wb.close()

    df = pd.DataFrame(rows, columns=["edition", "fy", "group", "product", "drug_name",
                                     "yakko_class", "unit", "sheet", "total_qty",
                                     "suppressed"])
    df.to_csv(os.path.join(OUT, "target_drugs_long.csv"), index=False, encoding="utf-8-sig")

    lst = pd.DataFrame(listing, columns=["fy", "sheet", "cls", "total"])
    thresholds = (lst.groupby(["fy", "sheet", "cls"])
                  .agg(n_listed=("total", "size"), min_listed_total=("total", "min"))
                  .reset_index())
    thresholds["cap"] = thresholds["n_listed"].map(listing_cap)
    thresholds["cap_hit"] = thresholds["cap"].notna()
    target_cls = df["yakko_class"].unique()
    thresholds = thresholds[thresholds["cls"].isin(target_cls)].reset_index(drop=True)
    thresholds.to_csv(os.path.join(OUT, "ndb_publication_thresholds.csv"), index=False,
                      encoding="utf-8-sig")

    bounds = publication_bounds(df, thresholds)
    bounds.to_csv(os.path.join(OUT, "censoring_bounds_long.csv"), index=False,
                  encoding="utf-8-sig")

    grp_ts = (df.groupby(["group", "fy"])["total_qty"].sum()
              .reset_index()
              .pivot(index="fy", columns="group", values="total_qty")
              .fillna(0.0))
    grp_ts.to_csv(os.path.join(OUT, "hcv_timeseries.csv"), encoding="utf-8-sig")

    prod_ts = (df.groupby(["group", "product", "fy"])["total_qty"].sum()
               .reset_index()
               .pivot_table(index=["group", "product"], columns="fy",
                            values="total_qty", fill_value=0.0))
    prod_ts.to_csv(os.path.join(OUT, "hcv_product_timeseries.csv"), encoding="utf-8-sig")

    # Upper-bound series: observed totals plus the maximum hidden by unpublished cells.
    add_prod = (bounds.groupby(["group", "product", "fy"])["max_hidden_qty"].sum()
                .unstack("fy").reindex(index=prod_ts.index, columns=prod_ts.columns)
                .fillna(0.0))
    prod_ts_upper = prod_ts + add_prod
    prod_ts_upper.to_csv(os.path.join(OUT, "hcv_product_timeseries_upper.csv"),
                         encoding="utf-8-sig")
    add_grp = (bounds.groupby(["fy", "group"])["max_hidden_qty"].sum().unstack("group")
               .reindex(index=grp_ts.index, columns=grp_ts.columns).fillna(0.0))
    (grp_ts + add_grp).to_csv(os.path.join(OUT, "hcv_timeseries_upper.csv"),
                              encoding="utf-8-sig")

    print("rows:", len(df), "suppressed cells:", int(df["suppressed"].sum()),
          "unpublished product-sheet-years bounded:", len(bounds))
    print(grp_ts.round(0))
    print("max hidden by group/fy:")
    print(add_grp.round(0))
    return df


if __name__ == "__main__":
    main()
