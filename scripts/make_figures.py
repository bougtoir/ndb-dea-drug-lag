#!/usr/bin/env python3
"""
Figures for the HCV interferon-to-DAA transition analysis.

Fig 1: Decline of interferon-based therapy after IFN-free DAAs became available,
       indexed to FY2014=100, with the total DAA dispensed quantity on a second
       axis and dated approval/reimbursement markers. Peginterferon and the
       first-generation protease inhibitors mark interferon-based therapy;
       conventional interferon is shown as a non-HCV-specific background series;
       ribavirin is shown separately because it accompanied both interferon-based
       therapy and the interferon-free sofosbuvir + ribavirin regimen.
Fig 2: Product-level DAA wave (stacked area) showing the FY2015 surge and
       subsequent decline.
Fig 3: Combined estimated national treatment volume (interferon-free DAA courses +
       peginterferon courses) and the DAA share of it, i.e. substitution and total
       volume on one common scale.

Language is controlled by --lang {ja,en}; all in-figure text switches accordingly.
"""
import argparse
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

DATA = os.path.join(os.path.dirname(__file__), "..", "data")
RES = os.path.join(os.path.dirname(__file__), "..", "results")
OUT = os.path.join(os.path.dirname(__file__), "..", "output")
os.makedirs(OUT, exist_ok=True)

# Dated announcement / reimbursement-listing events (NHI listing = 薬価収載).
# Placed at the fiscal year in which they occurred.
EVENTS = [
    (2014, "daclatasvir+asunaprevir NHI-listed\n(first interferon-free oral regimen, 2014-09)",
            "ダクラタスビル＋アスナプレビル収載\n（初の全経口IFNフリー, 2014-09）"),
    (2015, "sofosbuvir (2015-05) & ledipasvir/sofosbuvir (2015-08) NHI-listed",
            "ソホスブビル(2015-05)・レジパスビル/ソホスブビル(2015-08)収載"),
    (2017, "glecaprevir/pibrentasvir NHI-listed\n(pangenotypic, 2017-11)",
            "グレカプレビル/ピブレンタスビル収載\n（パンジェノ, 2017-11）"),
]

LAB = {
    "en": dict(
        title1="National dispensed quantity of hepatitis C antiviral drugs after\n"
               "interferon-free DAAs became available (Japan, NDB Open Data)",
        y1="Dispensed quantity, indexed to FY2014 = 100",
        y2="Interferon-free DAA dispensed quantity (units, millions)",
        peg="Peginterferon (syringes)",
        pi="Protease inhibitors given with peginterferon",
        conv="Conventional interferon (not hepatitis-C-specific)",
        rbv="Ribavirin (used with both regimens)",
        daa="Interferon-free DAA total",
        xlab="Fiscal year",
        title2="The interferon-free DAA wave: dispensed quantity by product (Japan, NDB Open Data)",
        y2b="Dispensed quantity (units, millions)",
        note="Metric: NDB national dispensed quantity. Units differ across products.",
        title3="Estimated national hepatitis C antiviral treatment volume and the\ninterferon-free DAA share of it (Japan, NDB Open Data)",
        y3="Estimated treatment courses per fiscal year",
        y3b="Interferon-free DAA share of estimated courses (%)",
        b_daa="Interferon-free DAA courses (estimated)",
        b_peg="Peginterferon courses (estimated)",
        l_share="DAA share of estimated courses",
        note3=("Estimates from documented regimen durations, not observed patient counts. "
               "Ribavirin is excluded because it accompanied both regimens."),
    ),
    "ja": dict(
        title1="IFNフリーDAA登場後のC型肝炎抗ウイルス薬の全国処方数量\n（日本, NDBオープンデータ）",
        y1="処方数量（FY2014=100 指数）",
        y2="IFNフリーDAA処方数量（百万単位）",
        peg="ペグインターフェロン（注射）",
        pi="ペグIFN併用のプロテアーゼ阻害薬",
        conv="従来型インターフェロン（C型肝炎専用でない）",
        rbv="リバビリン（旧新両レジメンで使用）",
        daa="IFNフリーDAA合計",
        xlab="年度",
        title2="IFNフリーDAAの波：製剤別処方数量（日本, NDBオープンデータ）",
        y2b="処方数量（百万単位）",
        note="指標：NDB処方数量。製剤により単位が異なる。",
        title3="推定全国C型肝炎抗ウイルス治療量とIFNフリーDAAの占める割合\n（日本, NDBオープンデータ）",
        y3="年度あたり推定治療コース数",
        y3b="推定コース数に占めるIFNフリーDAAの割合（%）",
        b_daa="IFNフリーDAAコース数（推定）",
        b_peg="ペグインターフェロンコース数（推定）",
        l_share="推定コース数に占めるDAAの割合",
        note3=("投与期間の文献情報に基づく推定値であり、観測された患者数ではない。"
               "リバビリンは旧新両レジメンで使用されたため除外。"),
    ),
}

PROD_LABEL = {
    "en": {
        "daclatasvir": "daclatasvir", "asunaprevir": "asunaprevir",
        "simeprevir": "simeprevir", "telaprevir": "telaprevir",
        "vaniprevir": "vaniprevir", "sofosbuvir": "sofosbuvir",
        "ledipasvir/sofosbuvir": "ledipasvir/sofosbuvir",
        "ombitasvir/paritaprevir/ritonavir": "ombitasvir/paritaprevir/\nritonavir",
        "elbasvir": "elbasvir", "grazoprevir": "grazoprevir",
        "glecaprevir/pibrentasvir": "glecaprevir/pibrentasvir",
        "sofosbuvir/velpatasvir": "sofosbuvir/velpatasvir",
    },
    "ja": {
        "daclatasvir": "ダクラタスビル", "asunaprevir": "アスナプレビル",
        "simeprevir": "シメプレビル", "telaprevir": "テラプレビル",
        "vaniprevir": "バニプレビル", "sofosbuvir": "ソホスブビル",
        "ledipasvir/sofosbuvir": "レジパスビル/ソホスブビル",
        "ombitasvir/paritaprevir/ritonavir": "オムビタスビル/パリタプレビル/\nリトナビル",
        "elbasvir": "エルバスビル", "grazoprevir": "グラゾプレビル",
        "glecaprevir/pibrentasvir": "グレカプレビル/ピブレンタスビル",
        "sofosbuvir/velpatasvir": "ソホスブビル/ベルパタスビル",
    },
}


def setup_font(lang):
    if lang == "ja":
        import matplotlib.font_manager as fm
        for path in ["/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
                     "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"]:
            if os.path.exists(path):
                fm.fontManager.addfont(path)
                matplotlib.rcParams["font.family"] = fm.FontProperties(fname=path).get_name()
                break
    matplotlib.rcParams["axes.unicode_minus"] = False


def fig1(lang):
    L = LAB[lang]
    ts = pd.read_csv(os.path.join(DATA, "hcv_timeseries.csv")).set_index("fy")
    y0 = int(ts.index.min())

    def idx(col):
        return ts[col] / ts[col].loc[y0] * 100

    daa_m = ts["DAA"] / 1e6

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(ts.index, idx("IFN_peg").values, "o-", color="#c0392b", lw=2.4, label=L["peg"])
    ax.plot(ts.index, idx("PI_ifn").values, "^-", color="#8e44ad", lw=1.8, label=L["pi"])
    ax.plot(ts.index, idx("IFN_conv").values, "v-", color="#7f8c8d", lw=1.6, label=L["conv"])
    ax.plot(ts.index, idx("ribavirin").values, "s--", color="#e67e22", lw=2.0, label=L["rbv"])
    ax.set_ylabel(L["y1"])
    ax.set_xlabel(L["xlab"])
    ax.set_ylim(0, 110)
    ax.set_title(L["title1"], fontsize=12)

    ax2 = ax.twinx()
    ax2.bar(daa_m.index, daa_m.values, color="#2980b9", alpha=0.25, label=L["daa"])
    ax2.set_ylabel(L["y2"])
    ax2.set_ylim(0, daa_m.max() * 1.25)

    for fy, en, ja in EVENTS:
        ax.axvline(fy, color="grey", ls=":", lw=1)
        ax.annotate((en if lang == "en" else ja), xy=(fy, 100),
                    xytext=(fy + 0.05, 62 - 11 * EVENTS.index((fy, en, ja))),
                    fontsize=7.2, color="#333",
                    arrowprops=dict(arrowstyle="-", color="grey", lw=0.6))

    lines, labels = ax.get_legend_handles_labels()
    l2, lb2 = ax2.get_legend_handles_labels()
    ax.legend(lines + l2, labels + lb2, loc="upper right", fontsize=9)
    ax.text(0.01, -0.13, L["note"], transform=ax.transAxes, fontsize=7, color="grey")
    fig.tight_layout()
    p = os.path.join(OUT, f"fig1_ifn_collapse_{lang}.png")
    fig.savefig(p, dpi=200)
    tiff = os.path.join(OUT, f"fig1_ifn_collapse_{lang}.tif")
    fig.savefig(tiff, dpi=300)
    print("wrote", p, tiff)
    plt.close(fig)


def fig2(lang):
    L = LAB[lang]
    prod = pd.read_csv(os.path.join(DATA, "hcv_product_timeseries.csv"))
    daa = prod[prod["group"] == "DAA"].drop(columns=["group"]).set_index("product")
    daa = daa / 1e6
    order = daa.sum(axis=1).sort_values(ascending=False).index.tolist()
    daa = daa.loc[order]
    years = [int(c) for c in daa.columns]

    fig, ax = plt.subplots(figsize=(10, 6))
    labels = [PROD_LABEL[lang].get(p, p) for p in daa.index]
    ax.stackplot(years, daa.values, labels=labels, alpha=0.9)
    ax.set_title(L["title2"], fontsize=12)
    ax.set_xlabel(L["xlab"])
    ax.set_ylabel(L["y2b"])
    ax.legend(loc="upper right", fontsize=7.5, ncol=2)
    ax.text(0.01, -0.13, L["note"], transform=ax.transAxes, fontsize=7, color="grey")
    fig.tight_layout()
    p = os.path.join(OUT, f"fig2_daa_wave_{lang}.png")
    fig.savefig(p, dpi=200)
    tiff = os.path.join(OUT, f"fig2_daa_wave_{lang}.tif")
    fig.savefig(tiff, dpi=300)
    print("wrote", p, tiff)
    plt.close(fig)


def fig3(lang):
    """Combined estimated treatment volume and the DAA share of it."""
    L = LAB[lang]
    with open(os.path.join(RES, "course_estimate.json"), encoding="utf-8") as f:
        ce = json.load(f)
    daa_c = {int(k): v for k, v in ce["estimated_courses_by_fy"].items()}
    peg_c = {int(k): v for k, v in ce["peginterferon"]["estimated_courses_by_fy"].items()}
    share = {int(k): v for k, v in
             ce["combined_treatment_volume"]["daa_share_by_fy"].items()}
    years = sorted(daa_c)

    fig, ax = plt.subplots(figsize=(10, 6))
    peg_v = [peg_c[y] for y in years]
    daa_v = [daa_c[y] for y in years]
    ax.bar(years, peg_v, color="#c0392b", label=L["b_peg"])
    ax.bar(years, daa_v, bottom=peg_v, color="#2980b9", label=L["b_daa"])
    ax.set_title(L["title3"], fontsize=12)
    ax.set_xlabel(L["xlab"])
    ax.set_ylabel(L["y3"])

    ax2 = ax.twinx()
    # a year with no estimated treatment has no share: leave a gap, not a zero
    share_v = [np.nan if share[y] is None else share[y] * 100 for y in years]
    ax2.plot(years, share_v, "o-", color="#2c3e50",
             lw=2.0, label=L["l_share"])
    ax2.set_ylabel(L["y3b"])
    ax2.set_ylim(0, 105)

    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper right", fontsize=9)
    ax.text(0.01, -0.13, L["note3"], transform=ax.transAxes, fontsize=7, color="grey")
    fig.tight_layout()
    p = os.path.join(OUT, f"fig3_treatment_volume_{lang}.png")
    fig.savefig(p, dpi=200)
    tiff = os.path.join(OUT, f"fig3_treatment_volume_{lang}.tif")
    fig.savefig(tiff, dpi=300)
    print("wrote", p, tiff)
    plt.close(fig)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--lang", choices=["ja", "en"], default="en")
    args = ap.parse_args()
    setup_font(args.lang)
    fig1(args.lang)
    fig2(args.lang)
    fig3(args.lang)
