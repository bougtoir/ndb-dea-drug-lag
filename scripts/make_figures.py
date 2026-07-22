#!/usr/bin/env python3
"""
Figures for the HCV new-treatment-anticipation / drug-lag analysis.

Fig 1: Collapse of the interferon-based "standard treatment" (peginterferon +
       ribavirin) after IFN-free DAAs became available, indexed to FY2014=100,
       with the total DAA dispensed quantity on a second axis, plus dated
       announcement markers ("報道側" events).
Fig 2: Product-level DAA wave (stacked area) showing the FY2015 surge and
       subsequent decline -- the signature of a finite pool of long-waiting
       patients ("待望論"), not a steady replacement flow.

Language is controlled by --lang {ja,en}; all in-figure text switches accordingly.
"""
import argparse
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

DATA = os.path.join(os.path.dirname(__file__), "..", "data")
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
        title1="Collapse of interferon-based standard therapy for hepatitis C\nafter interferon-free DAAs became available (Japan, NDB Open Data)",
        y1="IFN-based therapy, indexed to FY2014 = 100",
        y2="Interferon-free DAA dispensed quantity (units, millions)",
        peg="Peginterferon (syringes)", rbv="Ribavirin (capsules)",
        daa="Interferon-free DAA total",
        xlab="Fiscal year",
        title2="The interferon-free DAA wave: dispensed quantity by product (Japan, NDB Open Data)",
        y2b="Dispensed quantity (units, millions)",
        note="Metric: NDB national dispensed quantity. Units differ across products.",
    ),
    "ja": dict(
        title1="IFNフリーDAA登場後のC型肝炎インターフェロン標準治療の消失\n（日本, NDBオープンデータ）",
        y1="IFNベース治療（FY2014=100 指数）",
        y2="IFNフリーDAA処方数量（百万単位）",
        peg="ペグインターフェロン（注射）", rbv="リバビリン（カプセル）",
        daa="IFNフリーDAA合計",
        xlab="年度",
        title2="IFNフリーDAAの波：製剤別処方数量（日本, NDBオープンデータ）",
        y2b="処方数量（百万単位）",
        note="指標：NDB処方数量。製剤により単位が異なる。",
    ),
}

PROD_LABEL = {
    "en": {
        "daclatasvir": "daclatasvir", "asunaprevir": "asunaprevir",
        "simeprevir": "simeprevir", "telaprevir": "telaprevir",
        "vaniprevir": "vaniprevir", "sofosbuvir": "sofosbuvir",
        "ledipasvir/sofosbuvir": "ledipasvir/sofosbuvir",
        "ombitasvir/paritaprevir/ritonavir": "ombitasvir/paritaprevir/r",
        "elbasvir": "elbasvir", "grazoprevir": "grazoprevir",
        "glecaprevir/pibrentasvir": "glecaprevir/pibrentasvir",
        "sofosbuvir/velpatasvir": "sofosbuvir/velpatasvir",
    },
    "ja": {
        "daclatasvir": "ダクラタスビル", "asunaprevir": "アスナプレビル",
        "simeprevir": "シメプレビル", "telaprevir": "テラプレビル",
        "vaniprevir": "バニプレビル", "sofosbuvir": "ソホスブビル",
        "ledipasvir/sofosbuvir": "レジパスビル/ソホスブビル",
        "ombitasvir/paritaprevir/ritonavir": "オムビタスビル/パリタプレビル/r",
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
    peg_idx = ts["IFN_peg"] / ts["IFN_peg"].loc[2014] * 100
    rbv_idx = ts["ribavirin"] / ts["ribavirin"].loc[2014] * 100
    daa_m = ts["DAA"] / 1e6

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(peg_idx.index, peg_idx.values, "o-", color="#c0392b", lw=2.4, label=L["peg"])
    ax.plot(rbv_idx.index, rbv_idx.values, "s--", color="#e67e22", lw=2.0, label=L["rbv"])
    ax.set_ylabel(L["y1"])
    ax.set_xlabel(L["xlab"])
    ax.set_ylim(0, 110)
    ax.set_title(L["title1"], fontsize=12)

    ax2 = ax.twinx()
    ax2.bar(daa_m.index, daa_m.values, color="#2980b9", alpha=0.25, label=L["daa"])
    ax2.set_ylabel(L["y2"])
    ax2.set_ylim(0, daa_m.max() * 1.25)

    ei = 1 if lang == "en" else 2
    for fy, en, ja in EVENTS:
        ax.axvline(fy, color="grey", ls=":", lw=1)
        ax.annotate((en if lang == "en" else ja), xy=(fy, 100), xytext=(fy + 0.05, 70 - 12 * EVENTS.index((fy, en, ja))),
                    fontsize=7.2, color="#333",
                    arrowprops=dict(arrowstyle="-", color="grey", lw=0.6))

    lines, labels = ax.get_legend_handles_labels()
    l2, lb2 = ax2.get_legend_handles_labels()
    ax.legend(lines + l2, labels + lb2, loc="upper right", fontsize=9)
    ax.text(0.01, -0.13, L["note"], transform=ax.transAxes, fontsize=7, color="grey")
    fig.tight_layout()
    p = os.path.join(OUT, f"fig1_ifn_collapse_{lang}.png")
    fig.savefig(p, dpi=200)
    plt.close(fig)
    print("wrote", p)


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
    plt.close(fig)
    print("wrote", p)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--lang", choices=["ja", "en"], default="en")
    args = ap.parse_args()
    setup_font(args.lang)
    fig1(args.lang)
    fig2(args.lang)
