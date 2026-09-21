#!/usr/bin/env python3
"""
Generate the JA and EN manuscripts (docx), separate editable table docx, and
editable figure pptx for the HCV interferon-to-DAA transition study.

All result numbers are read from results/summary.json and data/*.csv and formatted
at runtime; no result value is hard-coded here. References are numbered in order of
first appearance (Vancouver). Tables are inserted immediately after the paragraph
that first cites them. For Hepatology Research and Emerging Infectious Diseases, figures are cited in the text but
supplied as separate files per each journal's Instructions for Authors; figure
legends and tables are collected after the references.

Usage:
    python3 scripts/make_manuscript.py
"""
import copy
import datetime
import json
import os
import re
import sys

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Inches, Pt, RGBColor, Mm
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt

BASE = os.path.join(os.path.dirname(__file__), "..")
DATA = os.path.join(BASE, "data")
RES = os.path.join(BASE, "results")
OUT = os.path.join(BASE, "output")
EID_COUNTS_PATH = os.path.join(OUT, "eid_word_counts.json")
JVH_COUNTS_PATH = os.path.join(OUT, "jvh_word_counts.json")
JGH_COUNTS_PATH = os.path.join(OUT, "jgh_word_counts.json")
JEPI_COUNTS_PATH = os.path.join(OUT, "jepi_word_counts.json")
JEPI_PAGEMAP_PATH = os.path.join(OUT, "jepi_page_map.json")

S = json.load(open(os.path.join(RES, "summary.json"), encoding="utf-8"))
ITS = json.load(open(os.path.join(RES, "its_summary.json"), encoding="utf-8"))
COURSE = json.load(open(os.path.join(RES, "course_estimate.json"), encoding="utf-8"))
TS = pd.read_csv(os.path.join(DATA, "hcv_timeseries.csv")).set_index("fy")
EV = pd.read_csv(os.path.join(DATA, "announcement_events.csv"))

# Literature-derived numbers used in the Introduction and Discussion.
Dsup = json.load(open(os.path.join(DATA, "discussion_support.json"), encoding="utf-8"))
IFN_COURSE = pd.read_csv(os.path.join(DATA, "ifn_course_assumptions.csv"))

Y0, Y1 = S["fiscal_year_range"]

# Access date for URLs in references, generated at build time.
_ACCESS_DATE_LONG = datetime.date.today().strftime("%d %B %Y")
_ACCESS_DATE_SHORT = datetime.date.today().strftime("%d %b %Y")
_ACCESS_DATE_JA = datetime.date.today().strftime("%Y年%m月%d日")

# Public repository (Data/Code Availability). Synced from the wip working repo.
REPO_URL = "https://github.com/bougtoir/ndb-dea-drug-lag"


def fmt(x, nd=0):
    return f"{x:,.{nd}f}"


def require(value, what):
    """Guard a value the manuscript prose describes explicitly.

    The Results and Discussion text names the fiscal year each series reached zero
    and the size of the ribavirin rebound. If a future data edition no longer shows
    that pattern, the numbers must not be silently formatted as 'FY-1' or crash
    mid-document: fail here with a message saying which sentence needs rewriting.
    """
    missing = value is None or (isinstance(value, int) and value == -1)
    if missing:
        raise SystemExit(
            f"make_manuscript: {what} is not present in the current results, but the "
            "manuscript text describes it. Update the corresponding sentences before "
            "regenerating.")
    return value


# ---- derived display values (all traceable to summary/timeseries) -------------
peg_drop = abs(S["peginterferon"]["pct_change_first_to_last"])
# First fiscal year from which the drug no longer appears among the published NDB
# cells (below the top-N listing cap or shown as '-'), not a true zero.
rbv_zero_year = require(S["ribavirin"]["first_year_below_publication_threshold"],
                        "the ribavirin below-threshold year")
rbv_low_year = S["ribavirin"]["first_year_near_zero"]
rbv_reb = require(S["ribavirin"]["rebound"], "the ribavirin apparent rebound")
require(rbv_reb["fold_increase"], "the ribavirin apparent rebound fold increase")
rbv_reb_artifact = S["ribavirin"]["rebound_within_censoring_bound"]
rbv_up_last = S["ribavirin"]["upper_bound_fy_last"]
pi_zero_year = require(S["protease_inhibitors_with_interferon"]["first_year_below_publication_threshold"],
                       "the protease-inhibitor below-threshold year")
pi_up_last = S["protease_inhibitors_with_interferon"]["upper_bound_fy_last"]
pi_first_up = S["protease_inhibitors_with_interferon"]["fy_first_upper_bound"]
CEN = S["censoring"]
cen_threshold = CEN["suppression_threshold"]
cen_n_suppressed = CEN["n_suppressed_target_cells"]
cen_n_rows = CEN["n_target_rows"]
cen_caps = sorted({int(v) for v in CEN["oral_antiviral_listing_cap_by_fy"].values()})
cen_daa_hidden_max = CEN["daa_total_pct_hidden_max_overall"]
daa_fall_up = abs(S["daa_total_upper_bound"]["pct_change_peak_to_last"])
TS_UP = pd.read_csv(os.path.join(DATA, "hcv_timeseries_upper.csv")).set_index("fy")
conv_drop = abs(S["interferon_conventional"]["pct_change_first_to_last"])
daa_peak_fy = S["daa_total"]["peak_fy"]
daa_peak_val_m = S["daa_total"]["peak_value"] / 1e6
daa_rise = S["daa_total"]["pct_change_first_to_peak"]
daa_fall = abs(S["daa_total"]["pct_change_peak_to_last"])
daa_last_m = S["daa_total"]["fy_last"] / 1e6
n_daa = S["n_distinct_daa_products"]

# ---- literature-derived display values (from data/discussion_support.json) ----
_DS = Dsup
CAR = _DS["carriers_2018"]
YAM = _DS["yamashita_2025"]
MMWR = _DS["mmwr_2022"]
DAA_RESP = _DS["dAA_phase3_response"]
ifn_weeks_min = int(IFN_COURSE["weeks_alt"].dropna().astype(int).min())
ifn_weeks_max = int(IFN_COURSE["weeks"].dropna().astype(int).max())
hcv_in_care_2011 = int(CAR["hcv_patients_in_care_2011"])
hcv_untreated_lo = int(CAR["hcv_diagnosed_untreated_2011"]["low"])
hcv_untreated_hi = int(CAR["hcv_diagnosed_untreated_2011"]["high"])
hcv_in_care_rounded = f"{hcv_in_care_2011 // 1000 * 1000:,}"
hcv_untreated_rounded = f"{hcv_untreated_lo // 1000 * 1000:,}-{hcv_untreated_hi // 1000 * 1000:,}"
hcv_in_care_wan = f"{round(hcv_in_care_2011 / 10000):.0f}"
hcv_untreated_wan_range = f"{hcv_untreated_lo / 10000:.1f}〜{hcv_untreated_hi / 10000:.1f}"
yam_coverage_pct = int(YAM["healthcare_coverage_pct"])
yam_patients = int(YAM["hcv_patients_treated_daa_2013_2022"])
yam_fall_pct = int(YAM["care_prevalence_fall_pct"])
yam_cost_peak_year = int(YAM["hepatitis_c_cost_peak_year"])
mmwr_pct_lo = min(MMWR["dAA_treatment_within_1_year_pct"].values())
mmwr_pct_hi = max(MMWR["dAA_treatment_within_1_year_pct"].values())
daa_cure_min_pct = int(DAA_RESP["cure_rate_pct_min"])
# Asia-Pacific elimination context (Journal of Gastroenterology and Hepatology variant).
POL = _DS["polaris_2022"]
RAZ = _DS["razavi_2020"]
GAM = _DS["gamkrelidze_2021"]
AUS = _DS["hajarizadeh_2018"]
TWN = _DS["chien_2021"]
pol_viraemic_m = float(POL["viraemic_infections_2020_millions"])
pol_treated_2020 = int(POL["treatment_initiations_2020"])
raz_n_countries = int(RAZ["high_income_countries_analyzed"])
raz_on_track = int(RAZ["countries_on_track_2030"])
raz_year = int(RAZ["assessment_year"])
gam_on_track = int(GAM["countries_on_track_2030"])
gam_year = int(GAM["assessment_year"])
aus_initiations_2016 = int(AUS["australia_daa_initiations_2016"])
aus_pct_2016 = int(AUS["australia_pct_of_chronic_hcv_2016"])
aus_start = AUS["program_start"]
twn_target = int(TWN["taiwan_treatment_target_patients"])
twn_target_year = int(TWN["taiwan_treatment_target_year"])
daa_dur_min = int(DAA_RESP["treatment_duration_weeks"]["low"])
daa_dur_max = int(DAA_RESP["treatment_duration_weeks"]["high"])


def decline_ci(d, ci_key="annual_change_hac95"):
    """Return (|annual %|, ci_low_magnitude, ci_high_magnitude) for a decay fit."""
    r = abs(d["annual_change_pct"])
    lo, hi = (abs(v) for v in d[ci_key])
    return r, min(lo, hi), max(lo, hi)


peg_r, peg_lo, peg_hi = decline_ci(ITS["peginterferon_decay"])
rbv_r, rbv_lo, rbv_hi = decline_ci(ITS["ribavirin_decay"])
conv_r, conv_lo, conv_hi = decline_ci(ITS["conventional_ifn_decay"])
rbv_fy0, rbv_fy1 = ITS["ribavirin_decay"]["fy_used"]
daa_knot = ITS["daa_segmented"]["knot_fy"]
daa_post_r = abs(ITS["daa_segmented"]["post_peak_annual_rate_pct"])
daa_post_ci = sorted(abs(v) for v in ITS["daa_segmented"]["post_peak_annual_rate_boot95"])
daa_pre_r = ITS["daa_segmented"]["pre_peak_annual_rate_pct"]
daa_pre_ci = sorted(ITS["daa_segmented"]["pre_peak_annual_rate_boot95"])
daa_pre_n = ITS["daa_segmented"]["n_obs_up_to_knot"]
daa_slope_p = ITS["daa_segmented"]["slope_change_p"]
n_obs = ITS["n_annual_observations"]
n_boot = ITS["bootstrap_resamples"]
rbv_up_years = ITS["ribavirin_decay"]["years_with_year_on_year_increase"]


def p_str(p):
    """P value shown as '<0.001' rather than a rounded zero."""
    return "P < 0.001" if p < 0.001 else f"P = {p:.3f}"


# treatment-course estimates (estimate only; see data/*_course_assumptions.csv)
course_peak_fy = COURSE["peak_fy"]
course_peak = COURSE["peak_estimated_courses"]
course_total_lo, course_total_hi = sorted(COURSE["estimated_total_courses_fy2014_2023_range"])
CV = COURSE["combined_treatment_volume"]
comb_peak_fy = CV["peak_fy"]
comb_peak = CV["peak_estimated_courses"]
comb_first = CV["fy_first_estimated_courses"]
comb_last = CV["fy_last_estimated_courses"]
comb_fall = abs(CV["pct_change_peak_to_last"])
share_first = require(CV["daa_share_by_fy"][str(Y0)], f"the FY{Y0} DAA share") * 100
share_knot = require(CV["daa_share_by_fy"][str(daa_peak_fy)],
                     f"the FY{daa_peak_fy} DAA share") * 100
share_last = require(CV["daa_share_by_fy"][str(Y1)], f"the FY{Y1} DAA share") * 100
peg_course_first = COURSE["peginterferon"]["estimated_courses_by_fy"][str(Y0)]
peg_course_last = COURSE["peginterferon"]["estimated_courses_by_fy"][str(Y1)]
peg_upc = COURSE["peginterferon"]["units_per_course"]
peg_upc_alt = COURSE["peginterferon"]["units_per_course_alt"]
comb_post_r = abs(ITS["combined_volume_segmented"]["post_peak_annual_rate_pct"])
comb_post_ci = sorted(
    abs(v) for v in ITS["combined_volume_segmented"]["post_peak_annual_rate_boot95"])
comb_slope_p = ITS["combined_volume_segmented"]["slope_change_p"]
_ITS_UP = ITS["censoring_upper_bound"]
daa_post_r_up = abs(_ITS_UP["daa_segmented"]["post_peak_annual_rate_pct"])
daa_post_ci_up = sorted(abs(v) for v in _ITS_UP["daa_segmented"]["post_peak_annual_rate_boot95"])
comb_post_r_up = abs(_ITS_UP["combined_volume_segmented"]["post_peak_annual_rate_pct"])
share_diff_max_pp = max(
    abs(v) for v in COURSE["censoring_upper_bound"]["daa_share_difference_vs_baseline_pp"].values())

# ------------------------------------------------------------------------------
# Reference text per source id (no fabricated citations; verifiable sources only).
# Citation NUMBERS are assigned dynamically in order of first appearance (Vancouver),
# tracked in CITE_ORDER during document build.
_REF_COMMON = {
    "tanaka_lag": "Tanaka M, Idei M, Sakaguchi H, Kato R, Sato D, Sawanobori K, et al. "
                  "Evolving landscape of new drug approval in Japan and lags from "
                  "international birth dates: retrospective regulatory analysis. "
                  "Clin Pharmacol Ther. 2021;109(5):1265-1273. https://doi.org/10.1002/cpt.2080",
    "kumada": "Kumada H, Suzuki Y, Ikeda K, Toyota J, Karino Y, Chayama K, et al. "
              "Daclatasvir plus asunaprevir for chronic HCV genotype 1b infection. "
              "Hepatology. 2014;59(6):2083-2091. https://doi.org/10.1002/hep.27113",
    "mizokami": "Mizokami M, Yokosuka O, Takehara T, Sakamoto N, Korenaga M, Mochizuki H, "
                "et al. Ledipasvir and sofosbuvir fixed-dose combination with and without "
                "ribavirin for 12 weeks in treatment-naive and previously treated "
                "Japanese patients with genotype 1 hepatitis C: an open-label, "
                "randomised, phase 3 trial. Lancet Infect Dis. 2015;15(6):645-653. "
                "https://doi.org/10.1016/S1473-3099(15)70099-X",
    "omata": "Omata M, Nishiguchi S, Ueno Y, Mochizuki H, Izumi N, Ikeda F, et al. "
             "Sofosbuvir plus ribavirin in Japanese patients with chronic genotype 2 "
             "HCV infection: an open-label, phase 3 trial. J Viral Hepat. "
             "2014;21(11):762-768. https://doi.org/10.1111/jvh.12312",
    "jsh": "Drafting Committee for Hepatitis Management Guidelines, the Japan Society of "
           "Hepatology. Japan Society of Hepatology guidelines for the management of "
           "hepatitis C virus infection: 2019 update. Hepatol Res. 2020;50(7):791-816. "
           "https://doi.org/10.1111/hepr.13503",
    "carriers": "Tanaka J, Akita T, Ohisa M, Sakamune K, Ko K, Uchida S, et al. Trends "
                "in the total numbers of HBV and HCV carriers in Japan from 2000 to "
                "2011. J Viral Hepat. 2018;25(4):363-372. https://doi.org/10.1111/jvh.12828",
    "setoyama": "Setoyama H, Tanaka Y, Kanto T. Seamless support from screening to "
                "anti-HCV treatment and HCC/decompensated cirrhosis: subsidy programs "
                "for HCV elimination. Glob Health Med. 2021;3(5):335-342. "
                "https://doi.org/10.35772/ghm.2021.01079",
    "yamashita": "Yamashita S, Kanda N, Hashimoto H, Yoshimoto H, Goda K, Mitsutake N, "
                 "et al. Trends in the healthcare burden of hepatitis C after the "
                 "introduction of direct-acting antivirals in Japan, 2013-2022: a "
                 "national claims database study in Japan. Int J Infect Dis. "
                 "2025;160:108043. https://doi.org/10.1016/j.ijid.2025.108043",
    "mmwr": "Thompson WW, Symum H, Sandul A, Gupta N, Patel P, Nelson N, et al. "
            "Vital signs: hepatitis C treatment among insured adults - United States, "
            "2019-2020. MMWR Morb Mortal Wkly Rep. 2022;71(32):1011-1017. "
            "https://doi.org/10.15585/mmwr.mm7132e1",
    "newey": "Newey WK, West KD. A simple, positive semi-definite, heteroskedasticity "
             "and autocorrelation consistent covariance matrix. Econometrica. "
             "1987;55(3):703-708. https://doi.org/10.2307/1913610",
    "efron": "Efron B, Tibshirani RJ. An Introduction to the Bootstrap. New York: "
             "Chapman & Hall; 1993.",
    # Asia-Pacific HCV elimination context (JGH variant); metadata verified against
    # Crossref/Europe PMC records at build time of this variant.
    "polaris": "Polaris Observatory HCV Collaborators. Global change in hepatitis C "
               "virus prevalence and cascade of care between 2015 and 2020: a modelling "
               "study. Lancet Gastroenterol Hepatol. 2022;7(5):396-415. "
               "https://doi.org/10.1016/S2468-1253(21)00472-6",
    "razavi": "Razavi H, Sanchez Gonzalez Y, Yuen C, Cornberg M. Global timing of "
              "hepatitis C virus elimination in high-income countries. Liver Int. "
              "2020;40(3):522-529. https://doi.org/10.1111/liv.14324",
    "gamkrelidze": "Gamkrelidze I, Pawlotsky JM, Lazarus JV, Feld JJ, Zeuzem S, Bao Y, "
                   "et al. Progress towards hepatitis C virus elimination in high-income "
                   "countries: an updated analysis. Liver Int. 2021;41(3):456-463. "
                   "https://doi.org/10.1111/liv.14779",
    "hajarizadeh": "Hajarizadeh B, Grebely J, Matthews GV, Martinello M, Dore GJ. Uptake "
                   "of direct-acting antiviral treatment for chronic hepatitis C in "
                   "Australia. J Viral Hepat. 2018;25(6):640-648. "
                   "https://doi.org/10.1111/jvh.12852",
    "apasl": "Omata M, Kanda T, Wei L, Yu ML, Chuang WL, Ibrahim A, et al. APASL "
             "consensus statements and recommendation on treatment of hepatitis C. "
             "Hepatol Int. 2016;10(5):702-726. https://doi.org/10.1007/s12072-016-9717-6",
    "tanaka_cm": "Tanaka J, Akita T, Ko K, Miura Y, Satake M; Epidemiological Research "
                 "Group on Viral Hepatitis and its Long-term Course, Ministry of Health, "
                 "Labour and Welfare of Japan. Countermeasures against viral hepatitis B "
                 "and C in Japan: an epidemiological point of view. Hepatol Res. "
                 "2019;49(9):990-1002. https://doi.org/10.1111/hepr.13417",
    "taiwan": "Chien RN, Lu SN, Pwu RF, Wu GH, Yang WW, Liu CL. Taiwan accelerates its "
              "efforts to eliminate hepatitis C. Glob Health Med. 2021;3(5):293-300. "
              "https://doi.org/10.35772/ghm.2021.01064",
}

REF_TEXT = {
    "en": dict(_REF_COMMON, **{
        "ndb": "Ministry of Health, Labour and Welfare (Japan). NDB Open Data "
               "(1st-10th editions). https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/"
               "0000177182.html "
               f"(accessed {_ACCESS_DATE_LONG}).",
        "bms": "Bristol-Myers Squibb K.K. Press release: approval in Japan of Daklinza "
               "(daclatasvir) and Sunvepra (asunaprevir), the world's first all-oral, "
               "interferon- and ribavirin-free treatment for chronic hepatitis C. "
               f"4 July 2014. https://www.bms.com/jp/media/press-release-listing/"
               "press-release-listing-2014/20140704.html "
               f"(accessed {_ACCESS_DATE_LONG}).",
        "nhi": "Ministry of Health, Labour and Welfare (Japan), Central Social Insurance "
               "Medical Council (Chuikyo). Records of National Health Insurance "
               "drug-price listings of direct-acting antivirals for hepatitis C, "
               "2014-2017 (Chuikyo plenary minutes and Health Insurance Bureau "
               "notifications; the document cited for each date is listed in the "
               f"study repository, {REPO_URL}). https://www.mhlw.go.jp/stf/shingi/shingi-chuo_128154.html "
               f"(accessed {_ACCESS_DATE_LONG}).",
        "subsidy": "Ministry of Health, Labour and Welfare (Japan). Special program for "
                   "the promotion of hepatitis treatment (medical-expense subsidy for "
                   "antiviral treatment of viral hepatitis). "
                   "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/kenkou/"
                   f"kekkaku-kansenshou/kanen/kangan/iryouhijyosei.html (accessed {_ACCESS_DATE_LONG}).",
    }),
    "ja": dict(_REF_COMMON, **{
        "ndb": "厚生労働省. NDBオープンデータ（第1〜10回）. "
               "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177182.html "
               f"（{Y0}〜{Y1}年度分、{_ACCESS_DATE_JA}アクセス）.",
        "bms": "ブリストル・マイヤーズ株式会社. プレスリリース：ダクルインザ錠"
               "（ダクラタスビル）およびスンベプラカプセル（アスナプレビル）"
               "—世界初のインターフェロン・リバビリンを必要としない経口薬のみによる"
               "C型慢性肝炎治療—の製造販売承認取得. 2014年7月4日. "
               "https://www.bms.com/jp/media/press-release-listing/"
               "press-release-listing-2014/20140704.html "
               f"（{_ACCESS_DATE_JA}アクセス）.",
        "nhi": "厚生労働省／中央社会保険医療協議会（中医協）. C型肝炎に対する直接作用型"
               "抗ウイルス薬の薬価基準収載記録（2014〜2017年；中医協総会資料および"
               "保険局医療課通知。日付ごとの根拠資料は data/announcement_events.csv に記載）. "
               f"https://www.mhlw.go.jp/stf/shingi/shingi-chuo_128154.html （{_ACCESS_DATE_JA}アクセス）.",
        "subsidy": "厚生労働省. 肝炎治療特別促進事業（ウイルス性肝炎の抗ウイルス治療に"
                   "対する医療費助成）. "
                   "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/kenkou/"
                   f"kekkaku-kansenshou/kanen/kangan/iryouhijyosei.html （{_ACCESS_DATE_JA}アクセス）.",
    }),
}

# EID follows the National Library of Medicine / Uniform Requirements style and asks
# that DOI URLs not be given for journal articles with a Medline link. References that
# are not journal articles (government data pages and press releases) keep their URLs.
_DOI_RE = re.compile(r"\s*https?://doi\.org/\S+")

# EID electronic citations use [cited DD Mon YYYY]. before the URL.
_EID_ACCESS_RE = re.compile(
    r'\s*(https?://\S+)\s*\(accessed\s+([^;)]+)(?:;\s*([^)]*))?\)\.'
)


def _format_eid_ref(text):
    text = _DOI_RE.sub("", text)

    def repl(m):
        url = m.group(1)
        date_raw = m.group(2).strip()
        extra = m.group(3)
        try:
            d = datetime.datetime.strptime(date_raw, "%d %B %Y")
            short = d.strftime("%d %b %Y")
        except ValueError:
            short = date_raw
        if extra:
            return f" [cited {short}]. {url} ({extra})."
        return f" [cited {short}]. {url}"

    return _EID_ACCESS_RE.sub(repl, text)


# JVH (Wiley) uses Vancouver/AMA reference style: <=6 authors listed, >=7 authors
# as first 3 + et al. (preceded by a comma), DOIs expressed as doi:..., and access
# dates placed after URLs.
_AMA_ACCESS_RE = re.compile(
    r'\s*(https?://\S+)\s*\(accessed\s+([^;)]*)(?:;\s*([^)]*))?\)\.'
)


def _format_jvh_ref(text):
    # Truncate long author lists (>=7 authors) to the first 3 + et al. for AMA style.
    # References with no "et al." (organisational authors) are left unchanged.
    if " et al." in text:
        before, after = text.split(" et al.", 1)
        author_blocks = [a.strip() for a in before.split(",") if a.strip()]
        if len(author_blocks) > 3:
            before = ", ".join(author_blocks[:3]) + ","
        text = before + " et al." + after
    # Prefer doi: prefix over https://doi.org/.
    text = re.sub(r'https?://doi\.org/([^ ]+)', r'doi:\1', text)

    def repl(m):
        url = m.group(1)
        date_raw = m.group(2).strip()
        extra = m.group(3)
        try:
            d = datetime.datetime.strptime(date_raw, "%d %B %Y")
            short = d.strftime("%B %d, %Y")
        except ValueError:
            short = date_raw
        if extra:
            return f" {url}. Accessed {short}. {extra}."
        return f" {url}. Accessed {short}."

    return _AMA_ACCESS_RE.sub(repl, text)


REF_TEXT_EID = {k: _format_eid_ref(v) for k, v in REF_TEXT["en"].items()}
REF_TEXT_JVH = {k: _format_jvh_ref(v) for k, v in REF_TEXT["en"].items()}
# JGH (Wiley) uses the same Vancouver conventions (all authors up to six, first
# three + et al. for seven or more, MEDLINE journal abbreviations).
REF_TEXT_JGH = REF_TEXT_JVH

CITE_ORDER = []  # reset per document; source ids in order of first appearance

# Default citation presentation; "hepres" uses bracketed numbers.
_CITATION_STYLE = "superscript"


def _set_citation_style(style):
    global _CITATION_STYLE
    _CITATION_STYLE = style


def _compress_citation_numbers(nums, sep=","):
    """Convert a sorted list of citation numbers to a compact range string."""
    if not nums:
        return ""
    parts = []
    start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
        else:
            parts.extend(_range_part(start, prev, sep))
            start = prev = n
    parts.extend(_range_part(start, prev, sep))
    return sep.join(parts)


def _range_part(start, prev, _sep=None):
    """Two adjacent numbers are written '1,2'; three or more as a range '1-3'."""
    if start == prev:
        return [str(start)]
    if prev == start + 1:
        return [str(start), str(prev)]
    return [f"{start}-{prev}"]


def cite(p, keys, lang):
    """Append a Vancouver citation, numbering by first appearance.
    Style is controlled by _CITATION_STYLE ('superscript', 'superscript_ama',
    'bracket', 'bracket_jvh', or 'paren_italic').
    """
    for k in keys:
        if k not in CITE_ORDER:
            CITE_ORDER.append(k)
    nums = sorted({CITE_ORDER.index(k) + 1 for k in keys})
    compressed = _compress_citation_numbers(nums)

    if _CITATION_STYLE == "bracket":
        run = p.add_run("[" + compressed + "]")
    elif _CITATION_STYLE == "bracket_jvh":
        # JVH places bracketed numbers after a word space, keeps the same
        # compressed ranges as Vancouver, and uses a comma+space separator
        # when multiple non-consecutive numbers are cited together.
        text = "[" + _compress_citation_numbers(nums, sep=", ") + "]"
        if p.runs and p.runs[-1].text and not p.runs[-1].text[-1].isspace():
            text = " " + text
        run = p.add_run(text)
    elif _CITATION_STYLE == "paren_italic":
        run = p.add_run("(" + compressed + ")")
        run.font.italic = True
    elif _CITATION_STYLE == "superscript_ama":
        run = p.add_run(_compress_citation_numbers(nums))
        run.font.superscript = True
    else:
        run = p.add_run(",".join(str(n) for n in nums))
        run.font.superscript = True
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(9)
    return p


def insert_fig(doc, path, width=6.3):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Inches(width))


# ---- native Word (OMML) equation helpers ------------------------------------
# In-text mathematics is written as genuine Word equations (Office Math markup),
# not as LaTeX source, so it renders and stays editable in Microsoft Word.

def _mr(text, italic=True, size='24'):
    """Return a native OMML math run (m:r) with Cambria Math font styling."""
    r = OxmlElement('m:r')
    # Upright function words (ln, max, annual change) need m:nor + m:sty p;
    # plain operators/digits are simply italic=False with no alpha characters.
    if not italic and any(ch.isalpha() for ch in text):
        rPr = OxmlElement('m:rPr')
        rPr.append(OxmlElement('m:nor'))
        sty = OxmlElement('m:sty')
        sty.set(qn('m:val'), 'p')
        rPr.append(sty)
        r.append(rPr)
    wRPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Cambria Math')
    rFonts.set(qn('w:hAnsi'), 'Cambria Math')
    wRPr.append(rFonts)
    if size:
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), size)
        wRPr.append(sz)
        szCs = OxmlElement('w:szCs')
        szCs.set(qn('w:val'), size)
        wRPr.append(szCs)
    r.append(wRPr)
    t = OxmlElement('m:t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    r.append(t)
    return r


def _msub(base, sub):
    """Return an OMML subscript (m:sSub). base and sub may be elements or lists."""
    s = OxmlElement('m:sSub')
    e = OxmlElement('m:e')
    for b in (base if isinstance(base, (list, tuple)) else [base]):
        e.append(b)
    sub_el = OxmlElement('m:sub')
    for sb in (sub if isinstance(sub, (list, tuple)) else [sub]):
        sub_el.append(sb)
    s.append(e)
    s.append(sub_el)
    return s


def _msup(base, sup):
    """Return an OMML superscript (m:sSup). base and sup may be elements or lists."""
    s = OxmlElement('m:sSup')
    e = OxmlElement('m:e')
    for b in (base if isinstance(base, (list, tuple)) else [base]):
        e.append(b)
    sup_el = OxmlElement('m:sup')
    for sp in (sup if isinstance(sup, (list, tuple)) else [sup]):
        sup_el.append(sp)
    s.append(e)
    s.append(sup_el)
    return s


def _append_omath(parent, items):
    """Append OMML elements (or nested lists of them) to a parent m:oMath.

    Elements are deep-copied so reusable tokens such as _Q_T can appear in
    multiple equations without being moved from one parent to another."""
    for item in (items if isinstance(items, (list, tuple)) else [items]):
        if isinstance(item, (list, tuple)):
            _append_omath(parent, item)
        else:
            parent.append(copy.deepcopy(item))


def add_equation(doc, elements, number=None):
    """Insert a centred, native Word equation from OMML element(s)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    math = OxmlElement('m:oMath')
    _append_omath(math, elements)
    para = OxmlElement('m:oMathPara')
    paraPr = OxmlElement('m:oMathParaPr')
    jc = OxmlElement('m:jc')
    jc.set(qn('m:val'), 'center')
    paraPr.append(jc)
    para.append(paraPr)
    para.append(math)
    p._p.append(para)
    if number:
        section = doc.sections[0]
        pos = section.page_width.inches - section.right_margin.inches - 0.1
        p.paragraph_format.tab_stops.add_tab_stop(Inches(max(pos, 0.5)), WD_TAB_ALIGNMENT.RIGHT)
        run = p.add_run(f"\t{number}")
        run.font.size = Pt(12)
    return p


def add_ieq(p, elements):
    """Append an inline native Word equation to an existing paragraph."""
    math = OxmlElement('m:oMath')
    _append_omath(math, elements)
    p._p.append(math)
    return p


# reusable inline math token factories (must return fresh elements so that
# copying a token into one equation does not remove it from another).
def _Q_T():
    return _msub(_mr('Q'), _mr('t'))


def _T_0():
    return _msub(_mr('t'), _mr('0'))


def _B1():
    return _msub(_mr('β'), _mr('1'))


def _B2():
    return _msub(_mr('β'), _mr('2'))


def _GAMMA():
    return _mr('γ')


def _KNOT():
    return [
        _msup([_mr('('), _mr('t'), _mr(' − '), _T_0(), _mr(')')], _mr('+')),
        _mr(' = ', italic=False),
        _mr('max', italic=False),
        _mr('(', italic=False),
        _mr('0', italic=False),
        _mr(', ', italic=False),
        _mr('t'),
        _mr(' − ', italic=False),
        _T_0(),
        _mr(')', italic=False),
    ]


def _ANN_B1():
    return [
        _msup(_mr('e'), _B1()),
        _mr(' − 1', italic=False),
    ]


def _ANN_B1B2():
    return [
        _msup(_mr('e'), [_mr('('), _B1(), _mr(' + '), _B2(), _mr(')')]),
        _mr(' − 1', italic=False),
    ]


def _ANN_G():
    return [
        _msup(_mr('e'), _GAMMA()),
        _mr(' − 1', italic=False),
    ]


# Segmented (broken-stick) log-linear model for total DAA dispensing, with a
# knot at the peak fiscal year t0:
def _EQ_SEGMENTED():
    return [
        _mr('ln ', italic=False),
        _Q_T(),
        _mr(' = ', italic=False),
        _msub(_mr('β'), _mr('0')),
        _mr(' + ', italic=False),
        _msub(_mr('β'), _mr('1')),
        _mr('(', italic=False),
        _mr('t'),
        _mr(' − ', italic=False),
        _T_0(),
        _mr(') + ', italic=False),
        _msub(_mr('β'), _mr('2')),
        _msup([_mr('('), _mr('t'), _mr(' − '), _msub(_mr('t'), _mr('0')), _mr(')')], _mr('+')),
        _mr(' + ', italic=False),
        _msub(_mr('ε'), _mr('t')),
    ]


# Exponential (log-linear) decay model for each interferon-based drug:
def _EQ_DECAY():
    return [
        _mr('ln ', italic=False),
        _Q_T(),
        _mr(' = ', italic=False),
        _mr('α'),
        _mr(' + ', italic=False),
        [_mr('γ'), _mr('t')],
        _mr(' + ', italic=False),
        _msub(_mr('ε'), _mr('t')),
        _mr(',', italic=False),
        _mr('    ', italic=False),
        _mr('annual change = ', italic=False),
        _msup(_mr('e'), _mr('γ')),
        _mr(' − 1', italic=False),
    ]


# ------------------------------------------------------------------------------
TXT = {
    "en": dict(
        title="How quickly did interferon-based therapy for hepatitis C give way to "
              "interferon-free direct-acting antivirals? A national dispensing analysis "
              f"of Japan's NDB Open Data, fiscal years {Y0}-{Y1}",
        h_abs="Abstract", h_intro="Introduction", h_meth="Methods",
        h_res="Results", h_disc="Discussion", h_lim="Limitations",
        h_conc="Conclusions",
        h_ref="References", h_da="Data and code availability",
        h_ack="Acknowledgments",
        abs_bg="Background: ", abs_me="Methods: ", abs_re="Results: ",
        abs_co="Conclusions: ",
    ),
    "ja": dict(
        title="C型肝炎のインターフェロン治療はIFNフリー直接作用型抗ウイルス薬へ"
              "どれだけ速く置換されたか：NDBオープンデータによる全国処方数量の解析"
              f"（{Y0}〜{Y1}年度）",
        h_abs="要旨", h_intro="緒言", h_meth="方法",
        h_res="結果", h_disc="考察", h_lim="限界", h_conc="結論",
        h_ref="文献", h_da="データおよびコードの入手可能性",
        abs_bg="背景：", abs_me="方法：", abs_re="結果：", abs_co="結論：",
    ),
}


def build_manuscript(lang, journal=None, inline=False):
    T = dict(TXT[lang])
    if journal == "pds":
        # Pharmacoepidemiology & Drug Safety structured-abstract heading.
        T["abs_bg"] = "Background/Objectives: "
    if journal == "hepres":
        # Hepatology Research: structured abstract with Aim/Methods/Results/Conclusions.
        T["title"] = (
            "Rapid population-level displacement of interferon-based hepatitis C therapy "
            "by interferon-free direct-acting antivirals in Japan"
        )
        T["abs_bg"] = "Aim: "
        T["h_res"] = "Results"
        T["h_disc"] = "Discussion"
        T["h_conc"] = "Conclusions"
    if journal == "eid":
        # Emerging Infectious Diseases (CDC): non-sentence title, unstructured abstract,
        # parenthetical italic numbered references, figures supplied as separate files.
        T["title"] = (
            "Rapid Population-Level Displacement of Interferon-Based Hepatitis C Therapy "
            "by Interferon-Free Direct-Acting Antivirals, Japan"
        )
        T["abs_bg"] = ""  # unstructured abstract; no subheadings
        T["h_res"] = "Results"
        T["h_disc"] = "Discussion"
        T["h_conc"] = "Conclusions"
    if journal == "jvh":
        # Journal of Viral Hepatitis (Wiley): title case, unstructured abstract,
        # bracketed Arabic citations in order of appearance, separate figure files,
        # and Vancouver/AMA-formatted references.
        T["title"] = (
            "Rapid Population-Level Displacement of Interferon-Based Hepatitis C Therapy "
            "by Interferon-Free Direct-Acting Antivirals in Japan"
        )
        T["abs_bg"] = ""
        T["abs_me"] = ""
        T["abs_re"] = ""
        T["abs_co"] = ""
        T["h_res"] = "Results"
        T["h_disc"] = "Discussion"
        T["h_conc"] = "Conclusions"
    if journal == "jgh":
        # Journal of Gastroenterology and Hepatology (Wiley): Original Article,
        # structured abstract (Background and Aim / Methods / Results / Conclusions),
        # 3-5 MeSH keywords, superscript Vancouver citations, 1.5 line spacing,
        # US spelling, figures as separate files, Introduction/Methods/Results/
        # Discussion headings with Conclusions inside the Discussion.
        T["title"] = (
            "Rapid nationwide transition from interferon-based to interferon-free "
            "hepatitis C therapy in Japan: treatment dynamics after reimbursement and "
            "their relevance to Asia-Pacific elimination programs"
        )
        T["abs_bg"] = "Background and Aim: "
        T["abs_me"] = "Methods: "
        T["abs_re"] = "Results: "
        T["abs_co"] = "Conclusions: "
        T["h_res"] = "Results"
        T["h_disc"] = "Discussion"
        T["h_conc"] = "Conclusions"
    if journal == "jepi":
        # Journal of Epidemiology (Japan Epidemiological Association, OA):
        # Original Article <=3,500 words; structured abstract
        # (Background/Methods/Results/Conclusions) <=250 words; 3-5 keywords;
        # superscript Vancouver (AMA) citations; sections Introduction, Methods,
        # Results, Discussion, Acknowledgments, Data Availability; continuous
        # line numbers from the Abstract through the Acknowledgments; tables as
        # editable text after the references; figures as separate files.
        T["title"] = (
            "Nationwide transition from interferon-based to interferon-free "
            "hepatitis C therapy in Japan: a descriptive study of national "
            "claims open data"
        )
        T["abs_bg"] = "Background: "
        T["abs_me"] = "Methods: "
        T["abs_re"] = "Results: "
        T["abs_co"] = "Conclusions: "
        T["h_res"] = "Results"
        T["h_disc"] = "Discussion"
        T["h_conc"] = "Conclusions"
    CITE_ORDER.clear()
    if journal == "hepres":
        _set_citation_style("bracket")
    elif journal == "eid":
        _set_citation_style("paren_italic")
    elif journal == "jvh":
        _set_citation_style("bracket_jvh")
    elif journal in ("jgh", "jepi"):
        _set_citation_style("superscript_ama")
    else:
        _set_citation_style("superscript")
    doc = Document()
    st = doc.styles["Normal"].font
    st.size = Pt(10.5)

    if journal == "hepres":
        _apply_hepres_format(doc)
    elif journal == "eid":
        _apply_eid_format(doc)
    elif journal == "jgh":
        _apply_jgh_format(doc)
    elif journal == "jepi":
        _apply_jepi_format(doc)

    if journal == "pds":
        ttl = doc.add_paragraph()
        r = ttl.add_run(T["title"]); r.bold = True; r.font.size = Pt(14)
        add_pds_titlepage(doc)
    elif journal == "hepres":
        add_hepres_titlepage(doc, T["title"])
    elif journal == "eid":
        add_eid_titlepage(doc, T["title"])
    elif journal == "jvh":
        add_jvh_titlepage(doc, T["title"])
    elif journal == "jgh":
        add_jgh_titlepage(doc, T["title"])
    elif journal == "jepi":
        add_jepi_titlepage(doc, T["title"])
    else:
        ttl = doc.add_paragraph()
        r = ttl.add_run(T["title"]); r.bold = True; r.font.size = Pt(14)

    # ---- Abstract ----
    doc.add_heading(T["h_abs"], level=1)
    if journal == "eid":
        _add_eid_abstract(doc, lang)
    elif journal == "jgh":
        _add_jgh_abstract(doc, T)
    elif journal == "jepi":
        _add_jepi_abstract(doc, T)
    else:
        ab = doc.add_paragraph()
        if lang == "en":
            ab.add_run(T["abs_bg"]).bold = True
            ab.add_run("Interferon-free direct-acting antivirals (DAAs) for hepatitis C virus "
                       "(HCV) were approved and covered by National Health Insurance (NHI) in "
                       "Japan in 2014-2015. How quickly an entire population moves away from an "
                       "established standard once a better option is approved and reimbursed is "
                       "rarely measured directly. We describe national antiviral use across "
                       "this transition. ")
            ab.add_run(T["abs_me"]).bold = True
            ab.add_run(f"Using the National Database of Health Insurance Claims and Specific "
                       f"Health Checkups of Japan (NDB) Open Data (fiscal years (FY) {Y0}-{Y1}), "
                       "we extracted national dispensed quantities of peginterferon, "
                       "first-generation non-structural protein 3/4A (NS3/4A) protease "
                       "inhibitors (given with peginterferon), conventional "
                       f"interferon, ribavirin, and {n_daa} interferon-free DAA products. "
                       "Quantities were converted to estimated treatment courses using "
                       "documented regimen durations, placing established and new therapy on a "
                       "common scale, and described with segmented log-linear and exponential trend "
                       "models. ")
            ab.add_run(T["abs_re"]).bold = True
            ab.add_run(f"Peginterferon dispensing fell {fmt(peg_drop,1)}% from FY{Y0} to "
                       f"FY{Y1} ({fmt(peg_r,0)}% per year, 95% confidence interval (CI) "
                       f"{fmt(peg_lo,0)}-{fmt(peg_hi,0)}%), and protease inhibitors fell below the "
                       f"NDB publication threshold from FY{pi_zero_year}. Estimated total treatment volume rose from "
                       f"{fmt(comb_first,0)} to a FY{comb_peak_fy} peak of {fmt(comb_peak,0)} "
                       f"courses, then fell {fmt(comb_fall,1)}% to {fmt(comb_last,0)} by "
                       f"FY{Y1}; the DAA share of estimated courses moved from "
                       f"{fmt(share_first,1)}% to {fmt(share_knot,1)}% within one fiscal year "
                       f"and stayed above {fmt(min(share_knot, share_last),0)}%. Ribavirin, "
                       "used in both interferon-based and interferon-free regimens, fell "
                       f"below the publication threshold from FY{rbv_zero_year}. ")
            ab.add_run(T["abs_co"]).bold = True
            ab.add_run("National hepatitis C treatment switched almost entirely to "
                       "interferon-free regimens within a year of reimbursement, with a "
                       "surge and decline in total volume of the kind expected when a "
                       "curative therapy works through a prevalent pool.")
        else:
            ab.add_run(T["abs_bg"]).bold = True
            ab.add_run("C型肝炎に対するIFNフリー直接作用型抗ウイルス薬（DAA）は2014〜2015年に"
                       "日本で承認・保険収載された。新たな選択肢が承認・償還された後、集団全体が"
                       "従来の標準治療からどれだけ速く離れるのかは直接測定されることが少ない。"
                       "本研究はこの移行期の全国的な抗ウイルス薬利用を記述する。")
            ab.add_run(T["abs_me"]).bold = True
            ab.add_run(f"NDBオープンデータ（{Y0}〜{Y1}年度）から、ペグインターフェロン、"
                       "第一世代NS3/4Aプロテアーゼ阻害薬（ペグインターフェロン併用）、従来型"
                       f"インターフェロン、リバビリン、および{n_daa}製剤のIFNフリーDAAの全国"
                       "処方数量を抽出した。数量は文献記載の投与期間を用いて推定治療コース数に"
                       "換算し、旧・新治療を同一尺度に置いた上で、分節対数線形モデルおよび指数"
                       "減衰モデルで記述した。")
            ab.add_run(T["abs_re"]).bold = True
            ab.add_run(f"ペグインターフェロンの処方数量はFY{Y0}からFY{Y1}で{fmt(peg_drop,1)}%"
                       f"減少し（年{fmt(peg_r,0)}%、95%CI {fmt(peg_lo,0)}〜{fmt(peg_hi,0)}%）、"
                       f"プロテアーゼ阻害薬はFY{pi_zero_year}以降公表下限未満となった。推定総治療量は"
                       f"{fmt(comb_first,0)}コースからFY{comb_peak_fy}の{fmt(comb_peak,0)}コースへ"
                       f"増加し、その後FY{Y1}までに{fmt(comb_fall,1)}%減少して{fmt(comb_last,0)}"
                       f"コースとなった。推定コース数に占めるDAAの割合は1年度で{fmt(share_first,1)}%"
                       f"から{fmt(share_knot,1)}%へ移行し、以後{fmt(min(share_knot, share_last),0)}%"
                       f"以上を維持した。旧・新両レジメンで用いられたリバビリンは"
                       f"FY{rbv_zero_year}以降公表下限未満となった。")
            ab.add_run(T["abs_co"]).bold = True
            ab.add_run("全国のC型肝炎治療は保険収載から1年以内にほぼ完全にIFNフリーレジメンへ"
                       "移行し、総治療量は治癒的治療が既存の有病者集団を処理する場合に予想される"
                       "急増と減衰を示した。")

    if journal == "pds":
        add_keywords(doc)
        add_key_points(doc)
        add_pls(doc)
    if journal == "hepres" and lang == "en":
        add_hepres_keywords(doc)
        add_hepres_abbreviations(doc)
    if journal == "jvh" and lang == "en":
        add_jvh_keywords(doc)
        add_jvh_abbreviations(doc)
    if journal == "jgh" and lang == "en":
        add_jgh_keywords(doc)
    if journal == "jepi" and lang == "en":
        add_jepi_keywords(doc)

    # ---- Introduction ----
    doc.add_heading(T["h_intro"], level=1)
    p = doc.add_paragraph()
    if journal == "jepi" and lang == "en":
        _add_jepi_introduction(doc, p, lang)
    elif journal == "jgh" and lang == "en":
        _add_jgh_introduction(doc, p, lang)
    elif lang == "en":
        p.add_run("Whether a newly approved therapy actually displaces the established "
                  "standard of care, and how fast, is a practical question for regulators "
                  "and payers as well as for patients waiting for a better option. Japan "
                  "has historically approved new drugs later than the United States or "
                  "Europe, although that gap narrowed over the 2010s")
        cite(p, ["tanaka_lag"], lang)
        p.add_run(", and approval is only the first step: a therapy also has to be "
                  "reimbursed, affordable, and taken up by clinicians and patients before "
                  "national utilization moves. Chronic hepatitis C provides an unusually "
                  "clear setting in which to observe that movement. Until 2014, treatment "
                  f"in Japan was built around interferon, given for {ifn_weeks_min}-{ifn_weeks_max} weeks with "
                  f"ribavirin and, from 2011, with a first-generation NS3/4A protease "
                  f"inhibitor; it was poorly tolerated and did not cure everyone")
        cite(p, ["jsh"], lang)
        p.add_run(". Interferon-free all-oral regimens then arrived in quick succession: "
                  "daclatasvir plus asunaprevir, approved in July 2014, was the first "
                  "regimen worldwide to dispense with both interferon and ribavirin")
        cite(p, ["bms", "kumada"], lang)
        p.add_run(", followed by sofosbuvir-based and pangenotypic regimens with cure "
                  f"rates above {daa_cure_min_pct}% and treatment durations of {daa_dur_min}-{daa_dur_max} weeks")
        cite(p, ["omata", "mizokami"], lang)
        p.add_run(".")

        if journal == "hepres" and lang == "en":
            p = doc.add_paragraph()
            p.add_run("This transition is also relevant to the global goal of hepatitis C "
                      "virus (HCV) elimination. The World Health Organization has set "
                      "targets for HCV elimination by 2030 through expanded testing and "
                      "treatment, but progress depends on whether health systems can move "
                      "quickly from older, less effective regimens to modern interferon-free "
                      "direct-acting antiviral therapy once it becomes available and "
                      "affordable")
            cite(p, ["setoyama"], lang)
            p.add_run(". Japan offers a national case study: it has a mature hepatitis "
                      "screening program and a subsidy program that caps out-of-pocket "
                      "costs for antiviral therapy, so the observed speed of the switch "
                      "reflects uptake under conditions designed to remove cost and access "
                      "barriers")
            cite(p, ["subsidy", "setoyama"], lang)
            p.add_run(". Examining how rapidly the country left interferon-based therapy "
                      "after interferon-free regimens were reimbursed can therefore inform "
                      "both national policy and the broader question of what is achievable "
                      "when curative treatment is listed promptly and subsidized.")

        p = doc.add_paragraph()
        p.add_run("Two outcomes are possible at the population level once such a therapy "
                  "is approved and reimbursed. National use of the prior standard may fall "
                  "steeply, indicating that the older therapy was being used for want of an "
                  "alternative and that little stands between reimbursement and treatment. "
                  "Alternatively, use may change slowly, which would point to real "
                  "practical friction—cost, referral pathways, capacity, or clinical "
                  "caution—between a therapy becoming available and a population receiving "
                  "it. Nationwide dispensing data can distinguish these, and the answer is "
                  "informative because the same question recurs with every new class of "
                  "therapy. We therefore used NDB Open Data, the national open-data release "
                  "of Japan's claims database")
        cite(p, ["ndb"], lang)
        p.add_run(", to describe how national dispensing of interferon-based therapy and of "
                  "interferon-free direct-acting antivirals (DAAs) changed around the approval and NHI-listing "
                  f"milestones of {Y0}-2017, how fast the two exchanged places on a common "
                  "estimated-course scale, and how much total hepatitis C treatment volume "
                  "the transition generated.")
    else:
        p.add_run("新たに承認された治療が実際に従来の標準治療をどれだけ速く置き換えるかは、"
                  "規制当局・支払者にとっても、より良い選択肢を待つ患者にとっても実務的な問いである。"
                  "日本の新薬承認は歴史的に米国・欧州より遅く、2010年代にその差は縮小したが")
        cite(p, ["tanaka_lag"], lang)
        p.add_run("、承認は第一段階にすぎない。全国的な利用が動くには、保険償還され、"
                  "費用負担が現実的で、臨床現場と患者に受容される必要がある。C型慢性肝炎は"
                  "その動きを観察するのに例外的に明瞭な場である。2014年までの日本の治療は"
                  f"インターフェロンを中心に構成され、リバビリンと{ifn_weeks_min}〜{ifn_weeks_max}週併用され、2011年以降は"
                  "第一世代NS3/4Aプロテアーゼ阻害薬も併用されたが、忍容性に乏しく全例を"
                  "治癒させるものではなかった")
        cite(p, ["jsh"], lang)
        p.add_run("。その後、IFNフリーの全経口レジメンが相次いで登場した。2014年7月に承認された"
                  "ダクラタスビル＋アスナプレビルは、インターフェロンとリバビリンの双方を不要と"
                  "した世界初のレジメンであり")
        cite(p, ["bms", "kumada"], lang)
        p.add_run(f"、続いてソホスブビル基盤およびパンジェノタイプのレジメンが、{daa_cure_min_pct}%を超える"
                  f"治癒率と{daa_dur_min}〜{daa_dur_max}週の投与期間で導入された")
        cite(p, ["omata", "mizokami"], lang)
        p.add_run("。")

        p = doc.add_paragraph()
        p.add_run("こうした治療が承認・収載された後、人口レベルでは二つの帰結が考えられる。"
                  "従来の標準治療の全国利用が急落する場合、旧治療は代替手段がないために"
                  "使われていたにすぎず、償還から治療到達までの障壁が小さいことを示す。"
                  "逆に変化が緩やかであれば、費用・紹介経路・診療能力・臨床的慎重さといった"
                  "実務的な摩擦が、利用可能性と実際の治療到達の間に存在することを示す。"
                  "全国処方データはこの二つを区別でき、同じ問いは新しい治療分類ごとに"
                  "繰り返されるため、その答えは有用である。そこで、日本のレセプト情報の"
                  "全国公開データであるNDBオープンデータを用い")
        cite(p, ["ndb"], lang)
        p.add_run(f"、{Y0}〜2017年の承認・薬価収載イベント前後で、インターフェロンベース治療と"
                  "IFNフリーDAAの全国処方数量がどう変化したか、両者が共通の推定コース尺度上で"
                  "どれだけ速く入れ替わったか、そして移行期に総治療量がどれだけ生じたかを"
                  "記述した。")

    # ---- Methods ----
    doc.add_heading(T["h_meth"], level=1)
    p = doc.add_paragraph()
    if lang == "en":
        p.add_run(f"Data. NDB Open Data editions 1-10 cover fiscal years {Y0}-{Y1}")
        cite(p, ["ndb"], lang)
        p.add_run(". From the sex- and age-stratified prescription-drug quantity tables "
                  "(oral, topical, injectable) we extracted the national total dispensed "
                  "quantity of every product of interest, identified by its marketed "
                  "product name. Dispensed quantity is counted in tablets or capsules for "
                  "oral drugs and in syringes or vials for injections; it is not a patient "
                  "count and is not comparable across products with different dosage units. "
                  "Japanese fiscal years run from April to March. In each formulation and "
                  "care-setting table, NDB Open Data lists only the highest-ranked products "
                  f"of each therapeutic class (top {cen_caps[0]} or {cen_caps[-1]}, by "
                  f"edition) and shows cells below {fmt(cen_threshold,0)} as a dash, so an "
                  "absent or dashed product is below the publication threshold, not "
                  "necessarily unused. Such cells are zero in the primary (lower-bound) "
                  "series; an upper-bound sensitivity series sets every "
                  "unpublished cell from a product's first listing onward to the largest "
                  "quantity it could conceal (the smallest published total of its class in "
                  f"that table when the listing cap was reached, otherwise "
                  f"{fmt(cen_threshold - 1,0)}). Of {cen_n_rows} extracted rows, "
                  f"{cen_n_suppressed} were dashed.")

        p = doc.add_paragraph()
        p.add_run("Drug groups. Peginterferon (peginterferon alfa-2a and alfa-2b) is the "
                  "hepatitis-C-specific marker of interferon-based therapy. The three "
                  "first-generation NS3/4A protease inhibitors (telaprevir, simeprevir, "
                  "vaniprevir) were always added to peginterferon plus ribavirin, so they "
                  "are treated as markers of interferon-based triple therapy and are kept "
                  "out of the DAA group. Conventional (non-pegylated) interferon is "
                  "reported separately as a background series because it is also used for "
                  "other indications and is not hepatitis-C-specific. Ribavirin is also "
                  "reported separately, and deliberately not as part of interferon-based "
                  "therapy, because it accompanied both peginterferon and the "
                  "interferon-free sofosbuvir plus ribavirin regimen for genotype 2")
        cite(p, ["omata"], lang)
        p.add_run(f"; its series therefore mixes established and new treatment. The remaining "
                  f"{n_daa} products form the interferon-free DAA group. Official approval "
                  "and NHI drug-price listing dates, taken from package inserts, Central "
                  "Social Insurance Medical Council (Chuikyo) records, and company "
                  "releases, are used as dated milestones at fiscal-year resolution "
                  "(Table 1)")
        cite(p, ["bms", "nhi"], lang)
        p.add_run(". No measure of news or social-media coverage, or of patient or "
                  "prescriber exposure to it, was available, so media effects were not "
                  "included as an exposure in this study.")
    else:
        p.add_run(f"データ。NDBオープンデータ第1〜10回は{Y0}〜{Y1}年度に対応する")
        cite(p, ["ndb"], lang)
        p.add_run("。処方薬の性年齢別薬効分類別数量表（内服・外用・注射）から、"
                  "対象製剤の全国総計処方数量を製品名により抽出した。処方数量は内服では"
                  "錠・カプセル、注射ではシリンジ・バイアルの数であり、患者数ではなく、"
                  "投与単位が異なる製剤間で比較できるものでもない。年度は4月〜3月である。"
                  "NDBオープンデータは剤形・診療区分ごとの表で薬効分類内の数量上位"
                  f"{cen_caps[0]}〜{cen_caps[-1]}品目のみを掲載し、{fmt(cen_threshold,0)}未満の"
                  "セルは「-」で表示する。したがって表に現れない、または「-」の製剤は"
                  "公表下限未満であって使用がないとは限らない。主解析（下限系列）では"
                  "これらを0とし、感度解析として、初掲載年度以降の非公表セルを隠れ得る"
                  "最大値（掲載上限に達した表ではその薬効分類の最小掲載値、それ以外は"
                  f"{fmt(cen_threshold - 1,0)}）で置き換えた上界系列を構成した。抽出した"
                  f"{cen_n_rows}行のうち{cen_n_suppressed}行で総計が「-」であった。")

        p = doc.add_paragraph()
        p.add_run("薬剤群。ペグインターフェロン（ペグインターフェロン アルファ-2a／2b）を"
                  "C型肝炎特異的なインターフェロンベース治療の指標とした。第一世代NS3/4A"
                  "プロテアーゼ阻害薬3剤（テラプレビル、シメプレビル、バニプレビル）は常に"
                  "ペグインターフェロン＋リバビリンに追加して用いられたため、インターフェロン"
                  "ベース3剤併用療法の指標として扱い、DAA群には含めない。従来型（非ペグ化）"
                  "インターフェロンは他疾患にも用いられC型肝炎特異的でないため、背景系列として"
                  "別掲した。リバビリンも別掲し、インターフェロンベース治療の一部としては"
                  "扱わない。リバビリンはペグインターフェロンと併用される一方、ジェノタイプ2に"
                  "対するIFNフリーのソホスブビル＋リバビリンでも用いられ")
        cite(p, ["omata"], lang)
        p.add_run(f"、その系列は旧・新治療を混在させるためである。残る{n_daa}製剤をIFNフリー"
                  "DAA群とした。添付文書、中央社会保険医療協議会（中医協）資料、企業"
                  "リリースから確認した承認日・薬価基準収載日を、年度解像度の日付付き"
                  "イベントとして用いた（表1）")
        cite(p, ["bms", "nhi"], lang)
        p.add_run("。報道量や、患者・処方医の報道接触を測る指標は"
                  "得られないため、本研究では報道量を曝露として扱わない。")

    if journal == "hepres" and lang == "en":
        p = doc.add_paragraph()
        p.add_run("During the preparation of this manuscript the authors used ChatGPT/GPT-4 for language editing and formatting assistance. After using this tool, the authors reviewed and edited the content and take full responsibility for the final manuscript.")

    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        add_caption(doc, ("Table 1. Official approval and NHI drug-price listing milestones "
                          "for interferon-free direct-acting antivirals in Japan."
                          if lang == "en"
                          else "表1．日本におけるIFNフリー直接作用型抗ウイルス薬の承認・"
                               "薬価基準収載イベント。"))
        add_events_table(doc, lang)

    p = doc.add_paragraph()
    if lang == "en":
        p.add_run("Common scale. Because dispensed quantities of an 8-week tablet regimen "
                  "and of a 48-week weekly injection are not comparable, we converted "
                  "quantities to estimated treatment courses using documented daily doses "
                  "and standard durations from Japanese package inserts and the Japan "
                  "Society of Hepatology (JSH) guideline")
        cite(p, ["jsh"], lang)
        p.add_run(". For each interferon-free regimen only one anchor component is counted, "
                  "so a two-drug regimen is not counted twice; for peginterferon one "
                  f"syringe is one weekly dose, giving {fmt(peg_upc,0)} syringes for a "
                  f"48-week course and {fmt(peg_upc_alt,0)} for a 24-week course as a "
                  "sensitivity assumption. Peginterferon courses also stand in for the "
                  "patients who additionally received a protease inhibitor, since those "
                  "drugs were always added to peginterferon; counting the inhibitors "
                  "separately would double-count the same treatment episodes. Ribavirin is "
                  "excluded from course estimates because it does not identify a single "
                  "regimen. Adding interferon-free DAA courses to peginterferon courses "
                  "gives an approximate national hepatitis C antiviral treatment volume, "
                  "and the DAA share of it measures substitution on one scale. These are "
                  "estimates that depend on the regimen assumptions, not observed patient "
                  "counts.")

        p = doc.add_paragraph()
        p.add_run(f"Statistical analysis. NDB Open Data begins in FY{Y0}, the period in "
                  "which interferon-free DAAs became available, so there is no "
                  "pre-intervention baseline and a conventional interrupted time series is "
                  f"not identifiable. With {n_obs} annual national observations we fitted "
                  "descriptive trend models. Total DAA dispensing and the combined "
                  "estimated treatment volume were each fitted with a continuous segmented "
                  "(broken-stick) log-linear model with a knot at the observed peak fiscal "
                  "year (Equation 1), giving separate pre- and post-peak annual "
                  "multiplicative rates. Each interferon-based drug was fitted with an "
                  "exponential (log-linear) decay over the fiscal years with positive "
                  "dispensing (Equation 2).")
        add_equation(doc, _EQ_SEGMENTED(), number="(1)")
        add_equation(doc, _EQ_DECAY(), number="(2)")
        pw = doc.add_paragraph()
        pw.add_run("where ")
        add_ieq(pw, _Q_T())
        pw.add_run(" is the national dispensed quantity in fiscal year ")
        add_ieq(pw, _mr("t"))
        pw.add_run(", ")
        add_ieq(pw, _T_0())
        pw.add_run(" the peak fiscal year, ")
        add_ieq(pw, _KNOT())
        pw.add_run("; ")
        add_ieq(pw, _B1())
        pw.add_run(" and ")
        add_ieq(pw, _B2())
        pw.add_run(" are the pre- and post-knot slope coefficients in Equation (1), ")
        add_ieq(pw, _GAMMA())
        pw.add_run(" is the slope in Equation (2), and the annual multiplicative change is ")
        add_ieq(pw, _ANN_B1())
        pw.add_run(" before the knot and ")
        add_ieq(pw, _ANN_B1B2())
        pw.add_run(" after the knot in Equation (1), and ")
        add_ieq(pw, _ANN_G())
        pw.add_run(" in Equation (2).")
        p2 = doc.add_paragraph()
        p2.add_run("Uncertainty is summarized with heteroskedasticity- and "
                   "autocorrelation-consistent (Newey-West, maxlags = 1) 95% intervals")
        cite(p2, ["newey"], lang)
        p2.add_run(f", cross-checked against a residual bootstrap ({fmt(n_boot,0)} "
                   "resamples)")
        cite(p2, ["efron"], lang)
        p2.add_run(". They describe how fast dispensing changed and how precisely that rate "
                   "can be pinned down; they are not causal estimates, and no comparison "
                   "condition is available. Because the DAA peak falls in the second "
                   f"fiscal year of the series, the pre-peak segment rests on {daa_pre_n} "
                   "observations and is reported for completeness rather than as an "
                   "estimated rate. Analyses used Python with statsmodels, and the pipeline "
                   "regenerates every reported number, figure, and table from the public "
                   "data.")
    else:
        p.add_run("共通尺度。8週の経口レジメンと48週の週1回注射では処方数量を比較できないため、"
                  "日本の添付文書および日本肝臓学会（JSH）ガイドライン記載の日用量・標準投与期間を"
                  "用いて、数量を推定治療コース数に換算した")
        cite(p, ["jsh"], lang)
        p.add_run("。各IFNフリーレジメンでは代表成分（アンカー）1つのみを計上し、2剤レジメンを"
                  "二重計上しない。ペグインターフェロンは1本＝1週分であり、48週コースは"
                  f"{fmt(peg_upc,0)}本、感度解析としての24週コースは{fmt(peg_upc_alt,0)}本とした。"
                  "プロテアーゼ阻害薬は常にペグインターフェロンに追加されたため、"
                  "ペグインターフェロンのコース数はその併用患者も含む。阻害薬を別に計上すれば"
                  "同一の治療エピソードを二重計上することになる。リバビリンは単一のレジメンを"
                  "特定しないためコース推定から除外した。IFNフリーDAAのコース数と"
                  "ペグインターフェロンのコース数を合算した値を、おおよその全国C型肝炎"
                  "抗ウイルス治療量とし、そのうちDAAが占める割合を同一尺度上の置換指標とした。"
                  "これらはレジメン仮定に依存する推定値であり、実測の患者数ではない。")

        p = doc.add_paragraph()
        p.add_run(f"統計解析。NDBオープンデータはFY{Y0}開始であり、これはIFNフリーDAAが"
                  "利用可能になった時期と重なるため、介入前の基準期間が存在せず、従来型の"
                  f"中断時系列は同定できない。{n_obs}点の年次全国データに記述的なトレンド"
                  "モデルを当てはめた。DAA合計処方数量および合算推定治療量は、それぞれ"
                  "観測ピーク年度をノットとする連続分節（折れ線）対数線形モデルで当てはめ（式1）、"
                  "ピーク前後の年次乗法的変化率を別々に得た。各インターフェロンベース薬は、"
                  "処方数量が正の年度に対する指数（対数線形）減衰で当てはめた（式2）。")
        add_equation(doc, _EQ_SEGMENTED(), number="(1)")
        add_equation(doc, _EQ_DECAY(), number="(2)")
        pw = doc.add_paragraph()
        pw.add_run("ここで ")
        add_ieq(pw, _Q_T())
        pw.add_run(" は年度 ")
        add_ieq(pw, _mr("t"))
        pw.add_run(" の全国処方数量、 ")
        add_ieq(pw, _T_0())
        pw.add_run(" はピーク年度、 ")
        add_ieq(pw, _KNOT())
        pw.add_run("；")
        add_ieq(pw, _B1())
        pw.add_run(" と ")
        add_ieq(pw, _B2())
        pw.add_run(" は式（1）のノット前後の傾き、")
        add_ieq(pw, _GAMMA())
        pw.add_run(" は式（2）の傾きであり、年次の乗法的変化は式（1）でノット前が ")
        add_ieq(pw, _ANN_B1())
        pw.add_run("、ノット後が ")
        add_ieq(pw, _ANN_B1B2())
        pw.add_run("、式（2）では ")
        add_ieq(pw, _ANN_G())
        pw.add_run(" で与えられる。")
        p2 = doc.add_paragraph()
        p2.add_run("不確実性は不均一分散・自己相関に頑健な（Newey-West, maxlags = 1）95%区間で表し")
        cite(p2, ["newey"], lang)
        p2.add_run(f"、残差ブートストラップ（{fmt(n_boot,0)}回）で相互確認した")
        cite(p2, ["efron"], lang)
        p2.add_run("。これらは変化の速さとその推定精度を記述するものであり、因果推定ではなく、"
                   "対照条件も存在しない。DAAのピークは系列の2年度目に位置するため、"
                   f"ピーク前区間は{daa_pre_n}点の観測に基づき、推定率としてではなく参考として"
                   "報告する。解析はPython（statsmodels）で行い、パイプラインは本文・図・表の"
                   "全数値を公開データから再生成する。")

    # ---- Results ----
    doc.add_heading(T["h_res"], level=1)
    p = doc.add_paragraph()
    if lang == "en":
        p.add_run("Interferon-based therapy fell away quickly (Figure 1, Table 2). "
                  f"Peginterferon dispensing dropped {fmt(peg_drop,1)}% between FY{Y0} and "
                  f"FY{Y1}, from {fmt(TS['IFN_peg'].loc[Y0],0)} to "
                  f"{fmt(TS['IFN_peg'].loc[Y1],0)} syringes, a decline of {fmt(peg_r,0)}% "
                  f"per year (95% CI {fmt(peg_lo,0)}-{fmt(peg_hi,0)}%). The "
                  "first-generation protease inhibitors, which were only ever given with "
                  f"peginterferon, went from {fmt(TS['PI_ifn'].loc[Y0],0)} published units "
                  f"in FY{Y0} to below the publication threshold in every table from FY{pi_zero_year} "
                  f"onward (at most {fmt(pi_up_last,0)} units in FY{Y1}). Conventional interferon, which is "
                  f"not hepatitis-C-specific, fell more gradually ({fmt(conv_drop,1)}% "
                  f"overall; {fmt(conv_r,0)}% per year, 95% CI {fmt(conv_lo,0)}-"
                  f"{fmt(conv_hi,0)}%), providing a background comparison; the much "
                  "steeper fall in peginterferon suggests that most of the "
                  "hepatitis-C-specific decline is not explained by a general drift in interferon use.")

        p = doc.add_paragraph()
        p.add_run("Ribavirin, shared by older and newer regimens, fell from "
                  f"{fmt(TS['ribavirin'].loc[Y0],0)} units in FY{Y0} to "
                  f"{fmt(rbv_reb['from_value'],0)} published units in FY{rbv_reb['from_fy']} "
                  f"and was below the publication threshold in every table from "
                  f"FY{rbv_zero_year} onward (at most {fmt(rbv_up_last,0)} units in FY{Y1}). "
                  f"The FY{rbv_reb['from_fy']}-FY{rbv_reb['to_fy']} values come only from "
                  "tables in which ribavirin remained listed, and their year-on-year "
                  "movement lies within what the unlisted tables could conceal, so it is "
                  "not interpreted. Fitted over the "
                  f"fiscal years with published dispensing (FY{rbv_fy0}-FY{rbv_fy1}), its "
                  f"decline was {fmt(rbv_r,0)}% per year "
                  f"(95% CI {fmt(rbv_lo,0)}-{fmt(rbv_hi,0)}%). Because ribavirin was used "
                  "with peginterferon and with interferon-free sofosbuvir, this rate "
                  "describes the combined retreat of both regimens and is not evidence "
                  "about interferon-based therapy alone.")
    else:
        p.add_run("インターフェロンベース治療は速やかに減少した（図1、表2）。"
                  f"ペグインターフェロンの処方数量はFY{Y0}の{fmt(TS['IFN_peg'].loc[Y0],0)}本から"
                  f"FY{Y1}の{fmt(TS['IFN_peg'].loc[Y1],0)}本へ{fmt(peg_drop,1)}%減少し、"
                  f"年{fmt(peg_r,0)}%（95%CI {fmt(peg_lo,0)}〜{fmt(peg_hi,0)}%）の減少率で"
                  "あった。ペグインターフェロンにのみ併用された第一世代プロテアーゼ阻害薬は、"
                  f"FY{Y0}の{fmt(TS['PI_ifn'].loc[Y0],0)}単位（非公表セルを考慮した上界"
                  f"{fmt(pi_first_up,0)}）からFY{pi_zero_year}以降は全ての表で公表下限未満と"
                  f"なった（FY{Y1}の上界{fmt(pi_up_last,0)}単位）。C型肝炎特異的でない従来型インターフェロンの減少はより緩やかで"
                  f"（全期間{fmt(conv_drop,1)}%、年{fmt(conv_r,0)}%、95%CI {fmt(conv_lo,0)}〜"
                  f"{fmt(conv_hi,0)}%）、C型肝炎特異的な減少のうちどれだけがインターフェロン"
                  "利用全般の背景的な減少で説明されうるかの目安となる。")

        p = doc.add_paragraph()
        p.add_run("旧・新両レジメンで共有されるリバビリンは、"
                  f"FY{Y0}の{fmt(TS['ribavirin'].loc[Y0],0)}単位からFY{rbv_reb['from_fy']}の"
                  f"公表値{fmt(rbv_reb['from_value'],0)}単位まで減少し、FY{rbv_zero_year}以降は"
                  f"全ての表で公表下限未満であった（FY{Y1}の上界{fmt(rbv_up_last,0)}単位）。"
                  f"FY{rbv_reb['from_fy']}〜FY{rbv_reb['to_fy']}の公表値は掲載順位を保った表のみに"
                  "由来し、その年次変動は非掲載の表が隠し得る範囲内にあるため解釈しない。"
                  f"公表値のある年度（FY{rbv_fy0}〜FY{rbv_fy1}）で"
                  f"当てはめた減少率は年{fmt(rbv_r,0)}%（95%CI {fmt(rbv_lo,0)}〜"
                  f"{fmt(rbv_hi,0)}%）であった。リバビリンはペグインターフェロンと、"
                  "またIFNフリーのソホスブビルと併用されたため、この率は両レジメンの"
                  "合わさった退潮を記述するものであり、インターフェロンベース治療のみに"
                  "関する証拠ではない。")
    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        insert_fig(doc, os.path.join(OUT, f"fig1_ifn_collapse_{lang}.png"))
    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        add_caption(doc, ("Figure 1. National dispensed quantity of hepatitis C antiviral "
                          f"drugs, indexed to FY{Y0} = 100, with total interferon-free DAA "
                          "dispensing on the right axis and dated approval and NHI-listing "
                          "milestones. Ribavirin and conventional interferon are shown "
                          "separately because they do not mark interferon-based hepatitis C "
                          "therapy uniquely." if lang == "en"
                          else f"図1．C型肝炎抗ウイルス薬の全国処方数量（FY{Y0}=100指数）。"
                               "右軸はIFNフリーDAA合計処方数量、縦線は承認・薬価収載イベント。"
                               "リバビリンと従来型インターフェロンはC型肝炎の"
                               "インターフェロンベース治療を一意に示さないため別系列とした。"))

    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        add_caption(doc, ("Table 2. National dispensed quantity by drug group and fiscal year "
                          "(NDB Open Data). Units differ across groups and are not additive. "
                          "Values are the published totals and are lower bounds; NP, not "
                          "published in any table (below the top-ranked listing cap of the "
                          f"class or shown as a dash for values below {fmt(cen_threshold,0)}), "
                          "with the maximum quantity the unpublished cells could conceal in "
                          "parentheses."
                          if lang == "en"
                          else "表2．薬剤群別・年度別の全国処方数量（NDBオープンデータ）。"
                               "群間で単位が異なり合算できない。値は公表値の合計であり下限値である。"
                               f"NP：いずれの表にも非公表（掲載上限外または{fmt(cen_threshold,0)}未満の"
                               "「-」表示）、括弧内は非公表セルが隠し得る最大量。"))
        add_quantity_table(doc, lang)

    p = doc.add_paragraph()
    if lang == "en":
        p.add_run("Interferon-free DAA dispensing rose and then receded (Figure 2). Total "
                  f"DAA dispensing peaked in FY{daa_peak_fy} at {fmt(daa_peak_val_m,1)} "
                  f"million units, {fmt(daa_rise,0)}% above FY{Y0}, and then fell "
                  f"{fmt(daa_fall,0)}% to {fmt(daa_last_m,1)} million units by FY{Y1}, with "
                  "each generation of product giving way to the next. In the segmented "
                  f"model, dispensing declined {fmt(daa_post_r,1)}% per year after the "
                  f"FY{daa_knot} peak (95% CI {fmt(daa_post_ci[0],1)}-"
                  f"{fmt(daa_post_ci[1],1)}%) and the post-peak slope differed from the "
                  f"pre-peak slope ({p_str(daa_slope_p)}). Because some later low-volume "
                  "cells are unpublished, the observed decline is an upper limit: in the "
                  "upper-bound series (unpublished "
                  f"cells at most {fmt(cen_daa_hidden_max,0)}% of the published DAA total in "
                  f"any year), the peak-to-FY{Y1} fall is {fmt(daa_fall_up,0)}% and the "
                  f"post-peak decline {fmt(daa_post_r_up,1)}% per year (95% CI "
                  f"{fmt(daa_post_ci_up[0],1)}-{fmt(daa_post_ci_up[1],1)}%). The pre-peak segment corresponds "
                  f"to {fmt(daa_pre_r,0)}% per year but rests on only {daa_pre_n} "
                  f"observations, and its bootstrap interval ({fmt(daa_pre_ci[0],1)}% to "
                  f"{fmt(daa_pre_ci[1],1)}%) includes zero; it is a single year-on-year "
                  "increment rather than an estimated trend and we do not interpret it as "
                  "a rate.")
    else:
        p.add_run("IFNフリーDAAの処方数量は増加し、その後退潮した（図2）。DAA合計は"
                  f"FY{daa_peak_fy}に{fmt(daa_peak_val_m,1)}百万単位（FY{Y0}比"
                  f"+{fmt(daa_rise,0)}%）でピークに達し、その後FY{Y1}までに{fmt(daa_fall,0)}%"
                  f"減少して{fmt(daa_last_m,1)}百万単位となり、世代ごとに製剤が入れ替わった。"
                  f"分節モデルでは、FY{daa_knot}のピーク後に年{fmt(daa_post_r,1)}%"
                  f"（95%CI {fmt(daa_post_ci[0],1)}〜{fmt(daa_post_ci[1],1)}%）で減少し、"
                  f"ピーク後の傾きはピーク前と異なった（{p_str(daa_slope_p)}）。"
                  f"ピーク前区間は年{fmt(daa_pre_r,0)}%に相当するが、{daa_pre_n}点の観測に"
                  f"基づくにすぎず、ブートストラップ区間（{fmt(daa_pre_ci[0],1)}%〜"
                  f"{fmt(daa_pre_ci[1],1)}%）はゼロを含む。これは単一の前年比増分であって"
                  "推定されたトレンドではなく、変化率として解釈しない。")
    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        insert_fig(doc, os.path.join(OUT, f"fig2_daa_wave_{lang}.png"))
    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        add_caption(doc, ("Figure 2. Interferon-free DAA dispensed quantity by product, "
                          "fiscal years " f"{Y0}-{Y1}." if lang == "en"
                          else f"図2．IFNフリーDAAの製剤別処方数量（{Y0}〜{Y1}年度）。"))

    p = doc.add_paragraph()
    if lang == "en":
        p.add_run("Placing both therapies on the estimated-course scale shows the exchange "
                  "directly (Figure 3, Table 3). Estimated peginterferon courses fell from "
                  f"about {fmt(peg_course_first,0)} in FY{Y0} to about "
                  f"{fmt(peg_course_last,0)} in FY{Y1}, while estimated interferon-free DAA "
                  f"courses peaked at about {fmt(course_peak,0)} in FY{course_peak_fy}. The "
                  f"combined estimated treatment volume rose from {fmt(comb_first,0)} "
                  f"courses in FY{Y0} to {fmt(comb_peak,0)} in FY{comb_peak_fy} and then "
                  f"declined {fmt(comb_fall,1)}% to {fmt(comb_last,0)} by FY{Y1} "
                  f"({fmt(comb_post_r,1)}% per year after the peak, 95% CI "
                  f"{fmt(comb_post_ci[0],1)}-{fmt(comb_post_ci[1],1)}%; slope change "
                  f"{p_str(comb_slope_p)}). The DAA share of estimated courses went from "
                  f"{fmt(share_first,1)}% in FY{Y0} to {fmt(share_knot,1)}% in "
                  f"FY{daa_peak_fy} and remained at {fmt(share_last,1)}% in FY{Y1}. Over "
                  f"FY{Y0}-FY{Y1} the estimated cumulative interferon-free DAA volume was "
                  f"{fmt(course_total_lo,0)}-{fmt(course_total_hi,0)} courses depending on "
                  "the duration assumption. These are estimates derived from regimen "
                  "durations, not counts of treated people.")
    else:
        p.add_run("両治療を推定コース尺度に置くと、入れ替わりが直接見える（図3、表3）。"
                  f"ペグインターフェロンの推定コース数はFY{Y0}の約{fmt(peg_course_first,0)}から"
                  f"FY{Y1}の約{fmt(peg_course_last,0)}へ減少し、IFNフリーDAAの推定コース数は"
                  f"FY{course_peak_fy}に約{fmt(course_peak,0)}でピークとなった。合算推定治療量は"
                  f"FY{Y0}の{fmt(comb_first,0)}コースからFY{comb_peak_fy}の{fmt(comb_peak,0)}"
                  f"コースへ増加し、その後FY{Y1}までに{fmt(comb_fall,1)}%減少して"
                  f"{fmt(comb_last,0)}コースとなった（ピーク後 年{fmt(comb_post_r,1)}%、"
                  f"95%CI {fmt(comb_post_ci[0],1)}〜{fmt(comb_post_ci[1],1)}%、傾きの変化"
                  f"{p_str(comb_slope_p)}）。推定コース数に占めるDAAの割合はFY{Y0}の"
                  f"{fmt(share_first,1)}%からFY{daa_peak_fy}に{fmt(share_knot,1)}%となり、"
                  f"FY{Y1}でも{fmt(share_last,1)}%であった。FY{Y0}〜FY{Y1}のIFNフリーDAAの"
                  f"推定累積治療量は、投与期間の仮定により{fmt(course_total_lo,0)}〜"
                  f"{fmt(course_total_hi,0)}コースであった。いずれも投与期間から導いた推定値で"
                  "あり、治療を受けた人数の計数ではない。")
    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        insert_fig(doc, os.path.join(OUT, f"fig3_treatment_volume_{lang}.png"))
    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        add_caption(doc, ("Figure 3. Estimated national hepatitis C antiviral treatment volume "
                          "(interferon-free DAA courses plus peginterferon courses) and the DAA "
                          "share of it. Ribavirin is excluded because it accompanied both "
                          "interferon-based and interferon-free regimens." if lang == "en"
                          else "図3．推定全国C型肝炎抗ウイルス治療量（IFNフリーDAAコース数＋"
                               "ペグインターフェロンコース数）とDAAの占める割合。リバビリンは"
                               "旧・新両レジメンで用いられたため除外。"))

    if journal not in ("hepres", "eid", "jvh", "jgh", "jepi") or inline:
        add_caption(doc, ("Table 3. Estimated treatment courses by fiscal year (estimates from "
                          "documented regimen durations, one anchor product per regimen; not "
                          "observed patient counts)." if lang == "en"
                          else "表3．年度別の推定治療コース数（文献記載の投与期間に基づく推定値、"
                               "レジメンごとに代表成分1つを計上。実測の患者数ではない）。"))
        add_course_table(doc, lang)

    # ---- Discussion ----
    doc.add_heading(T["h_disc"], level=1)
    p = doc.add_paragraph()
    if lang == "en":
        p.add_run("We asked whether national use of an established standard therapy falls "
                  "steeply once a new option is approved and reimbursed, or whether it "
                  "changes slowly enough to indicate practical friction. In hepatitis C the "
                  "first pattern is unambiguous. Dispensing shifted so fast that "
                  "interferon-free regimens accounted for almost all estimated treatment "
                  "courses within one fiscal year of the first interferon-free listing, and "
                  "the interferon-based drugs had fallen to negligible levels or below the "
                  "publication threshold within about two "
                  "years. This rapid displacement suggests that any friction between "
                  "reimbursement and treatment was small on an annual scale, although the "
                  "data do not isolate it.")

        p = doc.add_paragraph()
        p.add_run("The surge and decline in total treatment volume needs careful "
                  "interpretation. A curative therapy given to people who already have the "
                  "infection will produce exactly this shape without invoking anticipation "
                  "effects: the prevalent pool is treated over a few years and annual "
                  "demand then approaches the much smaller flow of newly diagnosed "
                  "patients. The pattern is therefore consistent with, but not evidence "
                  "for, a backlog of patients who had been waiting for a tolerable option. "
                  "Its magnitude does suggest that a large pool was available to treat: "
                  f"the estimated {fmt(course_total_lo,0)}-{fmt(course_total_hi,0)} "
                  "interferon-free DAA courses over ten years are of the same order as the "
                  f"hepatitis C population known to be in care in Japan before DAAs, "
                  f"estimated at about {hcv_in_care_rounded} patients receiving care in 2011 with a "
                  f"further {hcv_untreated_rounded} diagnosed but untreated")
        cite(p, ["carriers"], lang)
        p.add_run(f". A national claims study covering more than {yam_coverage_pct}% of Japanese healthcare "
                  f"reported {yam_patients:,} patients treated with DAAs between fiscal years 2013 "
                  f"and 2022, a {yam_fall_pct}% fall in the prevalence of care for chronic hepatitis C, "
                  f"and hepatitis C healthcare costs peaking in {yam_cost_peak_year}")
        cite(p, ["yamashita"], lang)
        p.add_run(". Our course estimates are somewhat lower, as expected when one anchor "
                  "component per regimen is converted with fixed durations, and the timing "
                  "of the peak agrees; together they suggest the dispensing-based estimates "
                  "track the treated population reasonably well.")

        p = doc.add_paragraph()
        p.add_run("Several factors besides a treatment backlog plausibly contributed to how "
                  "fast this happened, and the data cannot separate them. The new regimens "
                  f"were shorter, all-oral, and far better tolerated, with cure rates above "
                  f"{daa_cure_min_pct}% in Japanese phase 3 studies")
        cite(p, ["kumada", "omata", "mizokami", "jsh"], lang)
        p.add_run(", so both clinicians and patients had strong reasons to switch as soon as "
                  "the option existed. Japan also removes the cost barrier: on top of "
                  "universal insurance, the national program for the promotion of "
                  "hepatitis treatment subsidizes antiviral therapy so that most patients "
                  "pay a capped monthly amount, and hepatitis C DAAs were brought into that "
                  "program as they were listed")
        cite(p, ["subsidy", "setoyama"], lang)
        p.add_run(". Comparison with settings that lack such a scheme is instructive: among "
                  f"insured United States adults diagnosed in 2019-2020, only {mmwr_pct_lo}-{mmwr_pct_hi}% "
                  f"started DAA treatment within a year, varying by payer")
        cite(p, ["mmwr"], lang)
        p.add_run(". Availability alone therefore does not produce the Japanese pattern; the "
                  "observed speed is most likely the combined effect of a decisively "
                  "better therapy, prompt NHI listing, and subsidized affordability in a "
                  "population with a large diagnosed backlog. We measured none of the "
                  "mechanisms individually, and no measure of news coverage or of exposure "
                  "to it entered the analysis, so we make no claim about how information "
                  "reached patients or prescribers.")
        if journal == "jgh":
            _add_jgh_discussion_regional(doc, lang)
        if journal == "jepi":
            _add_jepi_discussion_opendata(doc, lang)
    else:
        p.add_run("本研究は、新たな選択肢が承認・償還された後に従来の標準治療の全国利用が"
                  "急落するのか、それとも実務的な摩擦を示すほど緩やかに変化するのかを問うた。"
                  "C型肝炎では前者が明瞭であった。処方は速やかに移行し、最初のIFNフリー"
                  "レジメン収載から1年度以内に推定治療コース数のほぼ全てがIFNフリーレジメンと"
                  "なり、インターフェロンベース薬は約2年で無視できる水準かゼロとなった。"
                  "本事例において償還と治療到達の間の摩擦は、年単位では小さかった。")

        p = doc.add_paragraph()
        p.add_run("総治療量の急増と減衰の解釈には注意が必要である。既に感染している人々に"
                  "治癒的治療を行えば、待望論を持ち出さなくとも同じ形状が生じる。すなわち"
                  "有病者プールが数年で治療され、その後の年間需要はより小さい新規診断の"
                  "流入に近づく。したがってこのパターンは、忍容可能な選択肢を待っていた"
                  "患者の滞留と整合的ではあるが、その証拠ではない。ただし規模は大きな"
                  f"プールの存在を示唆する。10年間の推定{fmt(course_total_lo,0)}〜"
                  f"{fmt(course_total_hi,0)}コースは、DAA前の日本で医療につながっていた"
                  f"C型肝炎人口（2011年に受療者約{hcv_in_care_wan}万人、診断済み未治療がさらに"
                  f"約{hcv_untreated_wan_range}万人と推計）と同程度の桁である")
        cite(p, ["carriers"], lang)
        p.add_run(f"。日本の医療の{yam_coverage_pct}%以上を捕捉する全国レセプト研究では、2013〜2022年度に"
                  f"{yam_patients:,}人がDAA治療を受け、C型慢性肝炎の受療割合は{yam_fall_pct}%減少し、"
                  f"C型肝炎関連医療費は{yam_cost_peak_year}年にピークを示したと報告されている")
        cite(p, ["yamashita"], lang)
        p.add_run("。本研究のコース推定はやや小さいが、これはレジメンごとに代表成分1つを"
                  "固定投与期間で換算した場合に予想されることであり、ピークの時期は一致する。"
                  "両者を合わせると、処方数量に基づく推定は治療人口をおおむね追跡できている"
                  "と考えられる。")

        p = doc.add_paragraph()
        p.add_run("この速さには治療の滞留以外の要因も寄与したと考えられ、本データでは分離できない。"
                  "新レジメンは投与期間が短く全経口で忍容性が大きく優れ、日本の第3相試験では"
                  f"{daa_cure_min_pct}%を超える治癒率が示された")
        cite(p, ["kumada", "omata", "mizokami", "jsh"], lang)
        p.add_run("ため、臨床医・患者の双方に選択肢が存在した時点で切り替える強い理由があった。"
                  "また日本では費用障壁が小さい。全国民保険に加え、肝炎治療特別促進事業が"
                  "抗ウイルス治療を助成し、多くの患者の自己負担は月額上限に抑えられ、"
                  "C型肝炎DAAも収載に合わせて対象とされた")
        cite(p, ["subsidy", "setoyama"], lang)
        p.add_run("。こうした制度を持たない環境との比較は示唆的である。2019〜2020年に"
                  "診断された米国の被保険成人では、1年以内にDAA治療を開始したのは支払者別に"
                  f"{mmwr_pct_lo}〜{mmwr_pct_hi}%にとどまった")
        cite(p, ["mmwr"], lang)
        p.add_run("。利用可能性のみでは日本のパターンは生じないため、ここで観察された速さは、"
                  "決定的に優れた治療、速やかな薬価収載、助成による費用負担の軽減が、"
                  "診断済みの大きな滞留人口を背景に重なった結果と考えられる。各機構を個別に"
                  "測定してはおらず、報道量やその接触の指標も解析に用いていないため、"
                  "情報が患者・処方医にどのように届いたかについては何も主張しない。")

    doc.add_heading(T["h_lim"], level=2)
    p = doc.add_paragraph()
    if lang == "en":
        p.add_run(f"NDB Open Data begins in FY{Y0}, the period in which interferon-free "
                  f"DAAs launched, so no pre-DAA interferon baseline exists within it and "
                  f"the FY{Y0} value already reflects decline from an earlier peak; the "
                  "analysis speaks to the speed of the observed displacement rather than "
                  "its full extent. Dispensed quantity is not a patient count, units are "
                  "not comparable across products, and the course figures are estimates "
                  "that inherit their regimen assumptions; the shorter-duration sensitivity "
                  "analysis indicates the size of that dependence. Because NDB Open Data "
                  "omits lower-ranked products and masks small cells, low-volume products "
                  "and the tail of a declining drug are underestimated, apparent "
                  "disappearance means only that a drug fell below the publication "
                  "threshold, and the later protease-inhibitor and ribavirin values are "
                  "lower bounds. The upper-bound analysis is a bound derived from the "
                  "publication rules, not an observed total; within it these omissions "
                  "change the post-peak decline of combined treatment volume from "
                  f"{fmt(comb_post_r,1)}% to {fmt(comb_post_r_up,1)}% per year and the "
                  f"interferon-free share by less than {fmt(share_diff_max_pp,1)} percentage "
                  "points, and alter neither the peak year nor the direction of any trend, "
                  "because the large early quantities that carry the main findings lie far "
                  "above the threshold. Ribavirin cannot be "
                  "assigned to either therapy, which is why it is reported separately and "
                  "excluded from course estimates, and conventional interferon is not "
                  "hepatitis-C-specific. The data are annual, so within-year timing around "
                  f"a listing date is invisible; with {n_obs} observations, no control "
                  "condition, and no pre-period, the trend models are descriptive and the "
                  "pre-peak DAA segment in particular rests on two observations. Secular "
                  "changes from FY2020, including the effect of the COVID-19 pandemic on "
                  "outpatient care, cannot be separated from the continuing decline. "
                  "Finally, no measure of media coverage or of exposure to it was "
                  "available, and the aggregate data cannot be stratified by genotype, "
                  "fibrosis stage, treatment history, or region, so nothing here identifies "
                  "why any individual was treated when they were.")
    else:
        p.add_run(f"NDBオープンデータはFY{Y0}開始であり、これはIFNフリーDAAの導入期と重なる"
                  f"ため、内部にDAA前のインターフェロン基準値は存在せず、FY{Y0}値は既に"
                  "それ以前のピークからの減少を反映している。本解析は観察された置換の速さを"
                  "示すものであり、その全体量を示すものではない。処方数量は患者数ではなく、"
                  "製剤間で単位は比較できず、コース数はレジメン仮定を引き継ぐ推定値である"
                  "（投与期間を短くした感度解析がその依存度を示す）。NDBオープンデータは各表で"
                  f"薬効分類の上位品目のみを掲載し{fmt(cen_threshold,0)}未満のセルを伏せるため、"
                  "低用量製剤と減衰する薬剤の末尾は過小評価され、見かけ上の消失は公表下限を"
                  f"下回ったことのみを意味し、FY{pi_zero_year}以降のプロテアーゼ阻害薬と"
                  f"FY{rbv_zero_year}以降のリバビリンの系列は下限値である。公表ルールから導いた"
                  "上界（観測総量ではない）による感度解析では、この欠落は単年でDAA合計の最大"
                  f"{fmt(cen_daa_hidden_max,0)}%に相当し、合算治療量のピーク後減少率を年"
                  f"{fmt(comb_post_r,1)}%から{fmt(comb_post_r_up,1)}%へ、IFNフリーの割合を"
                  f"{fmt(share_diff_max_pp,1)}ポイント未満しか変えず、ピーク年度もいずれの"
                  "傾向の方向も変えなかった。リバビリンはいずれの"
                  "治療にも帰属できないため別掲しコース推定から除外し、従来型インターフェロンは"
                  "C型肝炎特異的でない。データは年次であるため収載日前後の年内の時間構造は"
                  f"見えない。観測は{n_obs}点で対照条件も介入前期間もないため、トレンド"
                  "モデルは記述的であり、特にDAAのピーク前区間は2点に依拠する。FY2020以降は"
                  "COVID-19パンデミックの外来診療への影響を含む外生的変化が重なり、"
                  "継続的な減少と分離できない。さらに報道量やその接触の指標は得られず、"
                  "集計データはジェノタイプ・線維化進度・治療歴・地域で層別できないため、"
                  "個々の患者がその時期に治療された理由については何も同定しない。")

    # ---- Conclusions ----
    # JGH asks for Introduction/Methods/Results/Discussion as the main headings,
    # so Conclusions is a subsection of the Discussion there.
    # JE wants the concluding statement inside the Discussion (no separate
    # Conclusions section in its prescribed IMRaD+Acknowledgments+Data
    # Availability structure).
    if journal != "jepi":
        doc.add_heading(T["h_conc"], level=2 if journal == "jgh" else 1)
    p = doc.add_paragraph()
    if journal == "jepi" and lang == "en":
        p.add_run("In conclusion, national dispensing data show that hepatitis C "
                  "treatment in Japan moved almost entirely from interferon-based "
                  "to interferon-free regimens within one fiscal year of the first "
                  "interferon-free listing, and that total estimated treatment "
                  "volume surged then declined as the diagnosed prevalent pool was "
                  "treated. For health systems and for researchers using the same "
                  "open data, the episode demonstrates both what prompt listing "
                  "plus subsidy can achieve and how much of that trajectory can be "
                  "measured despite the NDB publication thresholds. Whether the "
                  "same speed is attainable for other therapies, or in systems "
                  "without comparable coverage and subsidy, cannot be inferred "
                  "from these data.")
    elif journal == "jgh" and lang == "en":
        p.add_run("National hepatitis C treatment in Japan moved almost entirely from "
                  "interferon-based to interferon-free regimens within a year of the first "
                  "interferon-free reimbursement, and interferon-based dispensing was "
                  "negligible within about two years. Total estimated treatment volume rose "
                  "and then fell steeply, as expected when a curative therapy works through "
                  "a prevalent population. For Asia-Pacific elimination programs the "
                  "trajectory is a planning template: once a decisively better therapy is "
                  "listed promptly and made affordable, the constraint moves within a few "
                  "years from treatment access to diagnosis and linkage to care. It does not "
                  "follow that the same therapy would be adopted as quickly in systems "
                  "without comparable coverage and subsidy.")
    elif lang == "en":
        p.add_run("National hepatitis C treatment in Japan moved almost entirely from "
                  "interferon-based to interferon-free regimens within a year of the first "
                  "interferon-free reimbursement, and interferon-based dispensing was "
                  "negligible within about two years. Total estimated treatment volume rose "
                  "and then fell steeply, as expected when a curative therapy works through "
                  "a prevalent population. The speed of the switch shows what is possible "
                  "when a decisively better therapy is listed promptly and made affordable; "
                  "it does not follow that other therapies, or the same therapy in systems "
                  "without comparable coverage and subsidy, would be adopted as quickly.")
    else:
        p.add_run("日本の全国的なC型肝炎治療は、最初のIFNフリーレジメンの保険償還から1年以内に"
                  "インターフェロンベースからIFNフリーへほぼ完全に移行し、インターフェロン"
                  "ベースの処方は約2年で無視できる水準となった。推定総治療量は増加後に急減し、"
                  "これは治癒的治療が有病者集団を処理する場合に予想される動きである。"
                  "この速さは、決定的に優れた治療が速やかに収載され費用負担が軽減された場合に"
                  "何が可能かを示すが、他の治療、あるいは同等の保険・助成制度を持たない"
                  "医療制度で同じ速さの導入が生じることを意味しない。")

    if journal == "pds":
        add_pds_statements(doc)

    if journal == "hepres" and lang == "en":
        add_hepres_declarations(doc, lang)
    elif journal == "eid" and lang == "en":
        add_eid_declarations(doc, lang)
    elif journal == "jvh" and lang == "en":
        add_jvh_declarations(doc, lang)
    elif journal == "jgh" and lang == "en":
        add_jgh_declarations(doc, lang)
    elif journal == "jepi" and lang == "en":
        add_jepi_backmatter(doc, lang)
    else:
        # ---- Data/code availability + references ----
        doc.add_heading(T["h_da"], level=1)
        p = doc.add_paragraph()
        p.add_run("All analysis code, derived datasets, and the figure/manuscript generators "
                  "are openly available at " + REPO_URL + ". The raw NDB Open Data workbooks "
                  "are published by the Ministry of Health, Labour and Welfare and are "
                  "re-downloadable with scripts/download_ndb.py, so the full pipeline "
                  "(download, build, analyze, figures, manuscript) reproduces every reported "
                  "number, figure, and table from the public data."
                  if lang == "en" else
                  "解析コード・派生データ・図表／原稿生成スクリプトはすべて " + REPO_URL + " で公開して"
                  "いる。NDBオープンデータの元ファイルは厚生労働省より公開され、scripts/download_ndb.py "
                  "で再取得できるため、全パイプライン（download→build→analyze→figures→manuscript）に"
                  "より本文・図・表の全数値を公開データから再現できる。")

    # JE line numbering runs from the Abstract through the Acknowledgments
    # only, so the references/tables/legends live in their own unnumbered
    # section. add_section clones the previous sectPr, so the inherited
    # lnNumType must be removed.
    if journal == "jepi" and lang == "en":
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        ln = sec._sectPr.find(qn("w:lnNumType"))
        if ln is not None:
            sec._sectPr.remove(ln)

    doc.add_heading(T["h_ref"], level=1)
    if journal == "eid" and lang == "en":
        ref_text = REF_TEXT_EID
    elif journal == "jvh" and lang == "en":
        ref_text = REF_TEXT_JVH
    elif journal in ("jgh", "jepi") and lang == "en":
        ref_text = REF_TEXT_JGH
    else:
        ref_text = REF_TEXT[lang]
    for i, k in enumerate(CITE_ORDER, 1):
        rp = doc.add_paragraph()
        rp.paragraph_format.left_indent = Inches(0.3)
        rp.paragraph_format.first_line_indent = Inches(-0.3)
        rp.add_run(f"{i}. {ref_text[k]}")

    if journal == "hepres" and lang == "en" and not inline:
        add_hepres_figure_legends(doc, lang)

    if journal == "eid" and lang == "en" and not inline:
        doc.add_page_break()
        add_tables_to_doc(doc, lang, heading=True)
        doc.add_page_break()
        add_eid_figure_legends(doc, lang)

    if journal == "jvh" and lang == "en" and not inline:
        doc.add_page_break()
        add_tables_to_doc(doc, lang, heading=True)
        doc.add_page_break()
        add_jvh_figure_legends(doc, lang)

    if journal == "jgh" and lang == "en" and not inline:
        doc.add_page_break()
        add_tables_to_doc(doc, lang, heading=True)
        doc.add_page_break()
        add_jgh_figure_legends(doc, lang)

    if journal == "jepi" and lang == "en" and not inline:
        doc.add_page_break()
        add_tables_to_doc(doc, lang, heading=True)
        doc.add_page_break()
        add_jgh_figure_legends(doc, lang)

    if journal == "hepres" and not inline:
        doc.add_page_break()
        add_tables_to_doc(doc, lang, heading=False)

    if journal in ("hepres", "eid", "jvh", "jgh", "jepi") and lang == "en":
        abs_words, main_words = _fill_word_counts(doc)
        if journal == "eid":
            counts = {
                "abstract_words": abs_words,
                "main_words": main_words,
                "reference_count": len(CITE_ORDER),
            }
            with open(EID_COUNTS_PATH, "w", encoding="utf-8") as fh:
                json.dump(counts, fh)
        if journal == "jvh":
            counts = {
                "abstract_words": abs_words,
                "main_words": main_words,
                "reference_count": len(CITE_ORDER),
            }
            with open(JVH_COUNTS_PATH, "w", encoding="utf-8") as fh:
                json.dump(counts, fh)
        if journal == "jgh":
            counts = {
                "abstract_words": abs_words,
                "main_words": main_words,
                "reference_count": len(CITE_ORDER),
            }
            with open(JGH_COUNTS_PATH, "w", encoding="utf-8") as fh:
                json.dump(counts, fh)
        if journal == "jepi":
            counts = {
                "abstract_words": abs_words,
                "main_words": main_words,
                "reference_count": len(CITE_ORDER),
            }
            with open(JEPI_COUNTS_PATH, "w", encoding="utf-8") as fh:
                json.dump(counts, fh)

    if journal and inline:
        suffix = f"_{journal}_inline"
    elif journal:
        suffix = f"_{journal}"
    else:
        suffix = ""
    path = os.path.join(OUT, f"manuscript_{lang}{suffix}.docx")
    doc.save(path)
    print("wrote", path)


def _fill_word_counts(doc):
    """Compute abstract and main-text word counts and fill title-page placeholders.

    Works for both Hepatology Research and Emerging Infectious Diseases layouts.
    """
    abs_words = 0
    main_words = 0
    in_abs = False
    in_main = False
    main_stop = (
        "References", "Acknowledgments", "Statements and Declarations",
        "Biographical Sketch", "Address for Correspondence",
        "Data and Code Availability", "Data Availability",
    )
    for p in doc.paragraphs:
        text = p.text.strip()
        if text == "Abstract":
            in_abs = True
            continue
        if in_abs and (text == "Introduction" or text.startswith("Keywords:")):
            in_abs = False
            if text == "Introduction":
                in_main = True
            continue
        if text == "Introduction":
            in_main = True
            continue
        if in_main and text in main_stop:
            in_main = False
            continue
        if in_abs:
            abs_words += len(text.split())
        if in_main and text and not (
            text.startswith("Figure") or text.startswith("Table")
        ):
            main_words += len(text.split())

    for p in doc.paragraphs:
        for r in p.runs:
            if "{{ABSTRACT_WORDS}}" in r.text:
                r.text = r.text.replace("{{ABSTRACT_WORDS}}", str(abs_words))
            if "{{MAIN_WORDS}}" in r.text:
                r.text = r.text.replace("{{MAIN_WORDS}}", str(main_words))
    return abs_words, main_words


def _set_no_vertical_borders(table):
    """Remove vertical table borders (Hepatology Research forbids vertical lines)."""
    tblPr = table._element.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        table._element.insert(0, tblPr)
    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)
    for edge in ("left", "right", "insideV"):
        edge_el = tblBorders.find(qn(f"w:{edge}"))
        if edge_el is None:
            edge_el = OxmlElement(f"w:{edge}")
            tblBorders.append(edge_el)
        edge_el.set(qn("w:val"), "none")
        edge_el.set(qn("w:sz"), "0")
        edge_el.set(qn("w:space"), "0")
        edge_el.set(qn("w:color"), "auto")


def _add_page_number_field(paragraph):
    """Insert a Word PAGE field into a paragraph (used for header page numbering)."""
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE'
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)


def _add_page_number_header(doc):
    """Add a right-aligned page number to the default header of every section."""
    for section in doc.sections:
        header = section.header
        p = header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.clear()  # remove any default content
        _add_page_number_field(p)


def _apply_hepres_format(doc):
    """Springer Nature Hepatology Research format: 10 pt Times Roman, double-spaced,
    30 mm margins, automatic page numbers."""
    for section in doc.sections:
        section.top_margin = Mm(30)
        section.bottom_margin = Mm(30)
        section.left_margin = Mm(30)
        section.right_margin = Mm(30)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(10)
    st.paragraph_format.line_spacing = 2.0
    _add_page_number_header(doc)


def add_hepres_titlepage(doc, title):
    """Springer Nature Hepatology Research title page (English)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True; r.font.size = Pt(14)
    p.paragraph_format.space_after = Pt(18)

    meta = [
        ("Article type", "Original Article"),
        ("Running head", "Interferon-to-DAA switch in Japan"),
        ("Authors", "[To be completed by the authors; include 16-digit ORCID iDs]"),
        ("Affiliations", "[To be completed by the authors; include institution, department, city, country]"),
        ("Corresponding author", "[Name], [Affiliation], [Full postal address], "
                                 "Tel: [telephone], Email: [email]"),
        ("Word count", "Abstract: {{ABSTRACT_WORDS}} words; main text: {{MAIN_WORDS}} words "
                       "(excluding references, declarations, tables, figure legends and acknowledgments)"),
        ("Tables / figures", "3 tables, 3 figures"),
        ("Acknowledgments", "None declared (people, grants, funds)."),
        ("Author contributions", "To be completed by the authors using the CRediT taxonomy. "
                                 "Example: [Author A]: Conceptualization, Methodology, "
                                 "Writing – original draft; [Author B]: ... All authors read and "
                                 "approved the final manuscript."),
    ]
    for k, v in meta:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)
    doc.add_page_break()


def add_hepres_keywords(doc):
    """Hepatology Research MeSH keywords, alphabetical, up to six."""
    p = doc.add_paragraph()
    p.add_run("Keywords: ").bold = True
    p.add_run("Antiviral Agents; Hepatitis C; Hepatitis C, Chronic; Interferons; "
              "Japan; Ribavirin")
    p.paragraph_format.space_after = Pt(12)


def add_hepres_abbreviations(doc):
    """Non-standard abbreviations used in the manuscript."""
    doc.add_heading("List of Abbreviations", level=2)
    items = [
        "CI: confidence interval",
        "DAA: direct-acting antiviral",
        "FY: fiscal year",
        "HCV: hepatitis C virus",
        "IFN: interferon",
        "JSH: Japan Society of Hepatology",
        "NDB: National Database of Health Insurance Claims and Specific Health Checkups of Japan",
        "NHI: National Health Insurance",
    ]
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(it)
    doc.add_paragraph()  # small space before Introduction


def add_hepres_declarations(doc, lang):
    """Springer Nature Hepatology Research Statements and Declarations section."""
    if lang != "en":
        return
    doc.add_heading("Statements and Declarations", level=1)
    items = [
        ("Funding", "This research received no specific grant from any funding agency."),
        ("Competing Interests", "The authors declare no competing interests."),
        ("Ethics approval", "This study used only publicly available, aggregated national open data "
                            "(NDB Open Data) containing no individual-level or identifiable information; "
                            "ethics-committee approval was not required."),
        ("Consent to participate", "Not applicable: this study used only aggregated, publicly "
                                    "available dispensing data."),
        ("Consent to publish", "Not applicable: this study used only aggregated, publicly "
                               "available dispensing data."),
        ("Data and code availability", "All analysis code, derived datasets, and the figure/manuscript "
                                       "generators are openly available at " + REPO_URL + ". The raw NDB "
                                       "Open Data workbooks are published by the Ministry of Health, Labour "
                                       "and Welfare and are re-downloadable with scripts/download_ndb.py, "
                                       "so the full pipeline (download, build, analyze, figures, manuscript) "
                                       "reproduces every reported number, figure, and table from the public data."),
        ("Authors' contributions", "To be completed by the authors using the CRediT taxonomy. "
                                   "Example: [Author A]: Conceptualization, Methodology, Writing – "
                                   "original draft; [Author B]: ... All authors read and approved the "
                                   "final manuscript."),
    ]
    for h, t in items:
        doc.add_heading(h, level=2)
        doc.add_paragraph(t)


def add_hepres_figure_legends(doc, lang):
    """Figure legends as a separate section after references (Hepatology Research)."""
    doc.add_heading("Figure Legends", level=1)
    legends = [
        ("Figure 1. National dispensed quantity of hepatitis C antiviral drugs, indexed to "
         f"FY{Y0} = 100, with total interferon-free DAA dispensing on the right axis and "
         "dated approval and NHI-listing milestones. Ribavirin and conventional interferon "
         "are shown separately because they do not mark interferon-based hepatitis C therapy "
         "uniquely."),
        (f"Figure 2. Interferon-free DAA dispensed quantity by product, fiscal years "
         f"{Y0}-{Y1}. Elbasvir and grazoprevir are shown separately because the NDB "
         "product files list the co-packaged components individually; ombitasvir/"
         "paritaprevir/ritonavir is a single co-formulated product."),
        ("Figure 3. Estimated national hepatitis C antiviral treatment volume "
         "(interferon-free DAA courses plus peginterferon courses) and the DAA share of it. "
         "Ribavirin is excluded because it accompanied both interferon-based and "
         "interferon-free regimens."),
    ] if lang == "en" else [
        (f"図1．C型肝炎抗ウイルス薬の全国処方数量（FY{Y0}=100指数）。右軸はIFNフリーDAA"
         "合計処方数量、縦線は承認・薬価収載イベント。リバビリンと従来型インターフェロンは"
         "C型肝炎のインターフェロンベース治療を一意に示さないため別系列とした。"),
        (f"図2．IFNフリーDAAの製剤別処方数量（{Y0}〜{Y1}年度）。エルバスビルとグラゾプレビルは"
         "配合錠の成分を個別に記載しているため別系列とした。オムビタスビル/パリタプレビル/"
         "リトナビルは単一配合剤である。"),
        ("図3．推定全国C型肝炎抗ウイルス治療量（IFNフリーDAAコース数＋ペグインターフェロン"
         "コース数）とDAAの占める割合。リバビリンは旧・新両レジメンで用いられたため除外。"),
    ]
    for lg in legends:
        p = doc.add_paragraph()
        p.add_run(lg)


def _set_line_numbering(section):
    """Enable continuous line numbering on a section."""
    sectPr = section._sectPr
    lnNumType = OxmlElement("w:lnNumType")
    lnNumType.set(qn("w:countBy"), "1")
    lnNumType.set(qn("w:start"), "1")
    lnNumType.set(qn("w:restart"), "continuous")
    sectPr.append(lnNumType)


def _apply_eid_format(doc):
    """Emerging Infectious Diseases format: 12 pt Times New Roman, double-spaced,
    left justified, 1-inch margins, and continuous line numbering."""
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.paragraph_format.line_spacing = 2.0
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_line_numbering(section)


def add_eid_titlepage(doc, title):
    """Emerging Infectious Diseases title page (English)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    p.paragraph_format.space_after = Pt(18)

    meta = [
        ("Article type", "Research"),
        ("Article summary line", "In Japan, hepatitis C antiviral use switched almost entirely to interferon-free direct-acting antiviral regimens within one fiscal year of reimbursement."),
        ("Running title", "Hepatitis C Treatment Switch in Japan"),
        ("Keywords", "Hepatitis C; Hepatitis C, Chronic; Antiviral Agents; Interferons; Direct-Acting Antivirals; Japan"),
        ("Authors", "[To be completed by the authors; include ORCID iDs for first and corresponding authors]"),
        ("Affiliations", "[To be completed by the authors; include institution, department, city, country]"),
        ("Corresponding author", "[Name], [Affiliation], [Full postal address], Phone: [telephone], Email: [email]"),
        ("Word count", "Abstract: {{ABSTRACT_WORDS}} words; main text: {{MAIN_WORDS}} words (excluding references, tables, figure legends, acknowledgments, and author information)"),
        ("Tables / figures", "3 tables, 3 figures"),
    ]
    for k, v in meta:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)
    doc.add_page_break()


def _add_eid_abstract(doc, lang):
    if lang != "en":
        return
    p = doc.add_paragraph()
    p.add_run("Interferon-free direct-acting antivirals (DAAs) for hepatitis C were "
              "reimbursed in Japan in 2014-2015. ")
    p.add_run(f"Using NDB Open Data (fiscal years {Y0}-{Y1}), we described national "
              "dispensed quantities of peginterferon, protease inhibitors, ribavirin, "
              "conventional interferon, and DAA products, converted to estimated treatment "
              "courses using documented regimen durations. ")
    p.add_run("Segmented log-linear and exponential trend models quantified the timing "
              "and rate of the interferon-to-DAA switch. ")
    p.add_run(f"Peginterferon dispensing fell {fmt(peg_drop,1)}% from FY{Y0} to FY{Y1}, "
              f"protease inhibitors fell below the NDB publication threshold from FY{pi_zero_year}, and estimated total "
              f"treatment volume peaked in FY{comb_peak_fy} and then fell "
              f"{fmt(comb_fall,1)}% to {fmt(comb_last,0)} courses by FY{Y1}. The DAA share "
              f"rose from {fmt(share_first,1)}% to {fmt(share_knot,1)}% within one fiscal "
              f"year and persisted above {fmt(min(share_knot, share_last),0)}%. ")
    p.add_run(f"Ribavirin fell below the publication threshold from FY{rbv_zero_year}. ")
    p.add_run("National hepatitis C treatment shifted almost entirely to interferon-free "
              "regimens within one year of reimbursement, with a surge-and-decline pattern "
              "consistent with a curative therapy moving through a prevalent pool.")


def add_eid_declarations(doc, lang):
    """Emerging Infectious Diseases back matter (English)."""
    if lang != "en":
        return
    doc.add_heading("Acknowledgments", level=1)
    p = doc.add_paragraph()
    p.add_run("This research received no specific grant from any funding agency. "
              "The authors declare no competing interests. The study used only publicly "
              "available, aggregated national open data (NDB Open Data) containing no "
              "individual-level or identifiable information; ethics-committee approval "
              "was not required. All analysis code, derived datasets, and the figure and "
              "manuscript generators are openly available at " + REPO_URL + ". The raw NDB "
              "Open Data workbooks are published by the Ministry of Health, Labour and "
              "Welfare and can be re-downloaded with scripts/download_ndb.py, so the full "
              "pipeline reproduces every reported number, figure, and table from the "
              "public data. During the preparation of this manuscript the authors used "
              "ChatGPT/GPT-4 for language editing and formatting assistance. After using "
              "this tool, the authors reviewed and edited the content and take full "
              "responsibility for the final manuscript.")
    doc.add_heading("Biographical Sketch", level=1)
    doc.add_paragraph("The first author is a physician-investigator and health services "
                      "researcher interested in the use of national claims and drug-utilization "
                      "data to evaluate how new treatments replace older standards of care. "
                      "Current work focuses on hepatitis C antiviral policy, pharmacoepidemiology, "
                      "and the translation of real-world dispensing data into public health insight.")
    doc.add_heading("Address for Correspondence", level=1)
    doc.add_paragraph("[To be completed by the authors: corresponding author name, full "
                      "postal address, email address, and telephone number.]")


def add_eid_figure_legends(doc, lang):
    """Figure legends as a separate section after tables (Emerging Infectious Diseases)."""
    if lang != "en":
        return
    doc.add_heading("Figure Legends", level=1)
    legends = [
        (f"Figure 1. National dispensed quantity of hepatitis C antiviral drugs, indexed to "
         f"FY{Y0} = 100, with total interferon-free DAA dispensing on the right axis and "
         "dated approval and NHI-listing milestones. Ribavirin and conventional interferon "
         "are shown separately because they do not mark interferon-based hepatitis C therapy "
         "uniquely."),
        (f"Figure 2. Interferon-free DAA dispensed quantity by product, fiscal years "
         f"{Y0}-{Y1}. Elbasvir and grazoprevir are shown separately because the NDB "
         "product files list the co-packaged components individually; ombitasvir/"
         "paritaprevir/ritonavir is a single co-formulated product."),
        ("Figure 3. Estimated national hepatitis C antiviral treatment volume "
         "(interferon-free DAA courses plus peginterferon courses) and the DAA share of it. "
         "Ribavirin is excluded because it accompanied both interferon-based and "
         "interferon-free regimens."),
    ]
    for lg in legends:
        p = doc.add_paragraph()
        p.add_run(lg)


def add_jvh_titlepage(doc, title):
    """Journal of Viral Hepatitis title page (English)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    p.paragraph_format.space_after = Pt(18)

    meta = [
        ("Article type", "Original Article"),
        ("Running title", "Interferon-to-DAA switch in Japan"),
        ("Authors", "[To be completed by the authors; include 16-digit ORCID iDs]"),
        ("Affiliations", "[To be completed by the authors; include institution, department, city, country]"),
        ("Corresponding author", "[Name], [Affiliation], [Full postal address], "
                                 "Tel: [telephone], Email: [email]"),
        ("Word count", "Abstract: {{ABSTRACT_WORDS}} words; main text: {{MAIN_WORDS}} words "
                       "(excluding references, declarations, tables, figure legends and acknowledgments)"),
        ("Tables / figures", "3 tables, 3 figures"),
        ("Acknowledgments", "None declared (people, grants, funds)."),
        ("Author contributions", "To be completed by the authors using the CRediT taxonomy. "
                                 "Example: [Author A]: Conceptualization, Methodology, "
                                 "Writing – original draft; [Author B]: ... All authors read and "
                                 "approved the final manuscript."),
    ]
    for k, v in meta:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)
    doc.add_page_break()


def add_jvh_keywords(doc):
    """Journal of Viral Hepatitis keywords (max 5)."""
    p = doc.add_paragraph()
    p.add_run("Keywords: ").bold = True
    p.add_run("Hepatitis C; Direct-Acting Antivirals; Interferons; Japan; NDB Open Data")
    p.paragraph_format.space_after = Pt(12)


def add_jvh_abbreviations(doc):
    """Non-standard abbreviations used in the manuscript."""
    doc.add_heading("List of Abbreviations", level=2)
    items = [
        "CI: confidence interval",
        "DAA: direct-acting antiviral",
        "FY: fiscal year",
        "HCV: hepatitis C virus",
        "JSH: Japan Society of Hepatology",
        "NDB: National Database of Health Insurance Claims and Specific Health Checkups of Japan",
        "NHI: National Health Insurance",
        "NS3/4A: non-structural protein 3/4A",
    ]
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(it)
    doc.add_paragraph()  # small space before Introduction


def add_jvh_declarations(doc, lang):
    """Journal of Viral Hepatitis Statements and Declarations section."""
    if lang != "en":
        return
    doc.add_heading("Statements and Declarations", level=1)
    items = [
        ("Funding", "This research received no specific grant from any funding agency."),
        ("Competing Interests", "The authors declare no competing interests."),
        ("Ethics approval", "This study used only publicly available, aggregated national open data "
                            "(NDB Open Data) containing no individual-level or identifiable information; "
                            "ethics-committee approval was not required."),
        ("Consent to participate", "Not applicable: this study used only aggregated, publicly "
                                    "available dispensing data."),
        ("Consent to publish", "Not applicable: this study used only aggregated, publicly "
                               "available dispensing data."),
        ("Data availability", "All analysis code, derived datasets, and the figure/manuscript "
                              "generators are openly available at " + REPO_URL + ". The raw NDB "
                              "Open Data workbooks are published by the Ministry of Health, Labour "
                              "and Welfare and are re-downloadable with scripts/download_ndb.py, "
                              "so the full pipeline (download, build, analyze, figures, manuscript) "
                              "reproduces every reported number, figure, and table from the public data."),
        ("Authors' contributions", "To be completed by the authors using the CRediT taxonomy. "
                                   "Example: [Author A]: Conceptualization, Methodology, Writing – "
                                   "original draft; [Author B]: ... All authors read and approved the "
                                   "final manuscript."),
    ]
    for h, t in items:
        doc.add_heading(h, level=2)
        doc.add_paragraph(t)


def add_jvh_figure_legends(doc, lang):
    """Figure legends as a separate section after tables (Journal of Viral Hepatitis)."""
    if lang != "en":
        return
    doc.add_heading("Figure Legends", level=1)
    legends = [
        (f"Figure 1. National dispensed quantity of hepatitis C antiviral drugs, indexed to "
         f"FY{Y0} = 100, with total interferon-free DAA dispensing on the right axis and "
         "dated approval and NHI-listing milestones. Ribavirin and conventional interferon "
         "are shown separately because they do not mark interferon-based hepatitis C therapy "
         "uniquely."),
        (f"Figure 2. Interferon-free DAA dispensed quantity by product, fiscal years "
         f"{Y0}-{Y1}. Elbasvir and grazoprevir are shown separately because the NDB "
         "product files list the co-packaged components individually; ombitasvir/"
         "paritaprevir/ritonavir is a single co-formulated product."),
        ("Figure 3. Estimated national hepatitis C antiviral treatment volume "
         "(interferon-free DAA courses plus peginterferon courses) and the DAA share of it. "
         "Ribavirin is excluded because it accompanied both interferon-based and "
         "interferon-free regimens."),
    ]
    for lg in legends:
        p = doc.add_paragraph()
        p.add_run(lg)


# ------------------------------------------------------------------------------
# Journal of Gastroenterology and Hepatology (Wiley) variant
# ------------------------------------------------------------------------------
def _apply_jgh_format(doc):
    """JGH: 12 pt Times New Roman, 1.5 line spacing, 25 mm margins, page numbers."""
    for section in doc.sections:
        section.top_margin = Mm(25)
        section.bottom_margin = Mm(25)
        section.left_margin = Mm(25)
        section.right_margin = Mm(25)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.paragraph_format.line_spacing = 1.5
    _add_page_number_header(doc)


def add_jgh_titlepage(doc, title):
    """JGH title page: title, running title, authors/affiliations, corresponding
    author, word counts, figure/table counts, and the title-page statements the
    journal asks for (conflict of interest, financial support, ethics, data)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    p.paragraph_format.space_after = Pt(18)

    meta = [
        ("Article type", "Original Article (Hepatology)"),
        ("Short title", "Interferon-to-DAA transition in Japan"),
        ("Authors", "[To be completed by the authors; include 16-digit ORCID iDs]"),
        ("Affiliations", "[To be completed by the authors; include institution, department, "
                         "city, country]"),
        ("Corresponding author", "[Name], [Affiliation], [Full postal address], "
                                 "Tel: [telephone], Email: [email]"),
        ("Word count", "Abstract: {{ABSTRACT_WORDS}} words; main text: {{MAIN_WORDS}} words "
                       "(Introduction to Discussion, excluding abstract, references, tables, "
                       "and figure legends)"),
        ("Tables and figures", "3 tables, 3 figures (6 display items in total)"),
        ("Declaration of conflict of interest", "The authors declare no conflict of interest."),
        ("Financial support", "This research received no specific grant from any funding "
                              "agency in the public, commercial, or not-for-profit sectors."),
        ("Ethics approval", "Not required: this study used only publicly available, "
                            "aggregated national open data (NDB Open Data) that contain no "
                            "individual-level or identifiable information."),
        ("Data availability", "All analysis code, derived datasets, and the figure/manuscript "
                              "generators are openly available at " + REPO_URL + "; the raw "
                              "NDB Open Data workbooks are published by the Ministry of "
                              "Health, Labour and Welfare and are re-downloadable with the "
                              "provided script."),
        ("Author contributions", "To be completed by the authors using the CRediT taxonomy. "
                                 "Example: [Author A]: Conceptualization, Methodology, "
                                 "Writing - original draft; [Author B]: ... All authors read "
                                 "and approved the final manuscript."),
    ]
    for k, v in meta:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)
    doc.add_page_break()


def _add_jgh_abstract(doc, T):
    """JGH structured abstract (Background and Aim / Methods / Results / Conclusions),
    <=250 words, written without abbreviations."""
    ab = doc.add_paragraph()
    ab.add_run(T["abs_bg"]).bold = True
    ab.add_run("Hepatitis C elimination depends on how quickly health systems move "
               "people to curative interferon-free therapy once it is reimbursed, yet "
               "national treatment dynamics after reimbursement are rarely measured. We "
               "describe the nationwide transition from interferon-based to "
               "interferon-free hepatitis C therapy in Japan, an Asia-Pacific country with "
               "universal insurance, a treatment subsidy program, and early listing of "
               "interferon-free regimens. ")
    ab.add_run(T["abs_me"]).bold = True
    ab.add_run(f"Using the national open-data release of Japan's health insurance claims "
               f"database (fiscal years {Y0}-{Y1}), we extracted national dispensed "
               "quantities of peginterferon, first-generation protease inhibitors given "
               "with peginterferon, conventional interferon, ribavirin, and "
               f"{_num_word(n_daa)} interferon-free direct-acting antiviral products. "
               "Quantities were converted to estimated treatment courses using documented "
               "regimen durations and described with segmented log-linear and exponential "
               "trend models. ")
    ab.add_run(T["abs_re"]).bold = True
    ab.add_run(f"Peginterferon dispensing fell {fmt(peg_drop,1)}% from fiscal year {Y0} to "
               f"{Y1} ({fmt(peg_r,0)}% per year, 95% confidence interval "
               f"{fmt(peg_lo,0)}-{fmt(peg_hi,0)}%), and protease inhibitors fell below the "
               f"publication threshold from fiscal year {pi_zero_year}. "
               "Estimated total treatment volume rose from "
               f"{fmt(comb_first,0)} to a fiscal year {comb_peak_fy} peak of "
               f"{fmt(comb_peak,0)} courses, then fell {fmt(comb_fall,1)}% to "
               f"{fmt(comb_last,0)} by fiscal year {Y1}; the interferon-free share of "
               f"estimated courses moved from {fmt(share_first,1)}% to "
               f"{fmt(share_knot,1)}% within one fiscal year and stayed above "
               f"{fmt(min(share_knot, share_last),0)}%. ")
    ab.add_run(T["abs_co"]).bold = True
    ab.add_run("Once interferon-free regimens were listed and subsidized, national "
               "hepatitis C treatment switched almost entirely within a year, and total "
               "volume surged and then declined as the prevalent pool was treated. The "
               "trajectory illustrates the dynamics Asia-Pacific elimination programs must "
               "plan for after reimbursement: rapid uptake, then falling demand that only "
               "case-finding and linkage to care can sustain.")


def _num_word(n):
    words = ["zero", "one", "two", "three", "four", "five", "six", "seven",
             "eight", "nine", "ten"]
    return words[n] if 0 <= n < len(words) else str(n)


def add_jgh_keywords(doc):
    """JGH keywords: 3-5 MeSH terms."""
    p = doc.add_paragraph()
    p.add_run("Keywords: ").bold = True
    p.add_run("Hepatitis C, Chronic; Antiviral Agents; Interferons; Disease Eradication; Japan")
    p.paragraph_format.space_after = Pt(12)


def _add_jgh_introduction(doc, p, lang):
    """JGH Introduction: Asia-Pacific HCV elimination context first, then Japan."""
    p.add_run("The World Health Organization has set targets to eliminate hepatitis C "
              "virus (HCV) infection as a public health threat by 2030, and treatment "
              "initiation is the step most directly under the control of health systems. "
              f"Globally, an estimated {pol_viraemic_m:.1f} million people had viremic "
              f"infection at the start of 2020 and about {pol_treated_2020:,} started "
              "treatment that year, a pace that is not on track to meet the targets")
    cite(p, ["polaris"], lang)
    p.add_run(". The Asia-Pacific region carries a large share of this burden, and "
              "although regional guidance has endorsed interferon-free direct-acting "
              "antivirals (DAAs) since 2016, access to them has been uneven across the "
              "region")
    cite(p, ["apasl", "polaris"], lang)
    p.add_run(". Where national programs have reimbursed DAAs broadly, uptake has been "
              f"rapid: Australia, which introduced government-funded DAA treatment "
              f"prescribable by specialists and general practitioners in {aus_start}, "
              f"treated an estimated {aus_initiations_2016:,} people, about "
              f"{aus_pct_2016}% of its chronically infected population, in the remainder of "
              "2016")
    cite(p, ["hajarizadeh"], lang)
    p.add_run(", and Taiwan combined public financing with relaxed reimbursement rules "
              f"toward a target of {twn_target:,} treated patients by {twn_target_year}")
    cite(p, ["taiwan"], lang)
    p.add_run(f". Modeling of {raz_n_countries} high-income countries placed Japan among "
              f"the {raz_on_track} on track for the 2030 targets at {raz_year} levels of "
              f"diagnosis and treatment and among the {gam_on_track} on track at {gam_year} "
              "levels, whereas the standing of some other Asia-Pacific economies changed "
              "between the two assessments")
    cite(p, ["razavi", "gamkrelidze"], lang)
    p.add_run(".")

    p = doc.add_paragraph()
    p.add_run("Japan's position rests on a mature national hepatitis policy that combines "
              "publicly funded hepatitis testing with a subsidy program that caps "
              "out-of-pocket payments for antiviral therapy on top of universal insurance")
    cite(p, ["tanaka_cm", "setoyama", "subsidy"], lang)
    p.add_run(". Until 2014, treatment was built around interferon, given for "
              f"{ifn_weeks_min}-{ifn_weeks_max} weeks with ribavirin and, from 2011, with a "
              "first-generation non-structural protein 3/4A (NS3/4A) protease inhibitor; "
              "it was poorly tolerated and did not cure everyone")
    cite(p, ["jsh"], lang)
    p.add_run(". Interferon-free all-oral regimens then arrived in quick succession: "
              "daclatasvir plus asunaprevir, approved in July 2014, was the first regimen "
              "worldwide to dispense with both interferon and ribavirin")
    cite(p, ["bms", "kumada"], lang)
    p.add_run(", followed by sofosbuvir-based and pangenotypic regimens with cure rates "
              f"above {daa_cure_min_pct}% and treatment durations of "
              f"{daa_dur_min}-{daa_dur_max} weeks")
    cite(p, ["omata", "mizokami"], lang)
    p.add_run(". Japan has historically approved new drugs later than the United States "
              "or Europe")
    cite(p, ["tanaka_lag"], lang)
    p.add_run(", but each DAA was listed on National Health Insurance (NHI) within months "
              "of approval and brought into the subsidy program as it was listed")
    cite(p, ["nhi", "subsidy"], lang)
    p.add_run(".")

    p = doc.add_paragraph()
    p.add_run("What such a setting produces at the population level has rarely been "
              "measured directly, yet it matters for elimination planning. If "
              "reimbursement removes the main barrier, national use of the older therapy "
              "should collapse and total treatment volume should surge as the diagnosed "
              "prevalent pool is treated, then fall toward the flow of new diagnoses; if "
              "practical friction—referral pathways, treatment capacity, or clinical "
              "caution—dominates, the switch should be gradual. The shape of that "
              "trajectory determines how much treatment capacity a program must mobilize "
              "and how soon its binding constraint shifts from treatment to case-finding. "
              "We therefore used NDB Open Data, the national open-data release of Japan's "
              "claims database")
    cite(p, ["ndb"], lang)
    p.add_run(", to describe how national dispensing of interferon-based therapy and of "
              "interferon-free DAAs changed around the approval and NHI-listing milestones "
              f"of {Y0}-2017, how fast the two exchanged places on a common "
              "estimated-course scale, and how much total hepatitis C treatment volume the "
              "transition generated.")


def _add_jgh_discussion_regional(doc, lang):
    """JGH Discussion paragraph: what the Japanese trajectory implies for
    Asia-Pacific elimination programs. Claims are limited to what the data and
    the cited literature support."""
    p = doc.add_paragraph()
    p.add_run("The Japanese trajectory bears directly on elimination planning elsewhere in "
              "the Asia-Pacific region. Australia, which began government-funded DAA "
              f"treatment in {aus_start}, likewise recorded a large initial wave of "
              "treatment initiation")
    cite(p, ["hajarizadeh"], lang)
    p.add_run(", and Taiwan built its national program around a time-limited treatment "
              "target backed by public financing and progressively relaxed reimbursement "
              "rules")
    cite(p, ["taiwan"], lang)
    p.add_run(". In each case the operational question is the one our data address: after "
              "reimbursement, how quickly is the diagnosed pool treated, and how soon does "
              "the binding constraint move from treatment access to diagnosis and linkage "
              "to care? In Japan the switch itself took about one fiscal year and the bulk "
              f"of the volume was dispensed within roughly five years, after which "
              f"estimated courses declined by about {fmt(comb_post_r,0)}% per year. "
              "Sustaining progress in that later phase depends on case-finding rather than "
              "on treatment supply: in Japan the groups now explicitly targeted by national "
              "countermeasures are undiagnosed carriers and diagnosed carriers who have "
              "never consulted, or have dropped out of, medical care")
    cite(p, ["tanaka_cm"], lang)
    p.add_run(". Elimination forecasts that hold diagnosis and treatment at recent levels")
    cite(p, ["razavi", "gamkrelidze"], lang)
    p.add_run(" should therefore be read against where a country sits on this curve: a "
              "falling treatment count after a completed switch is the expected signature "
              "of a treated pool, not necessarily of a failing program. Conversely, where "
              "DAAs are approved but reimbursement is restricted or out-of-pocket costs "
              "remain high")
    cite(p, ["polaris", "mmwr"], lang)
    p.add_run(", the same therapy will not produce this pattern. Japan's experience shows "
              "what prompt listing combined with subsidy can achieve in a population with "
              "a large diagnosed pool; it does not show what approval alone achieves, and "
              "it does not identify which of the contributing mechanisms mattered most.")


def add_jgh_declarations(doc, lang):
    """JGH end-of-text acknowledgments (declarations proper sit on the title page)."""
    if lang != "en":
        return
    doc.add_heading("Acknowledgments", level=1)
    doc.add_paragraph("None declared (people, grants, funds).")


def add_jgh_figure_legends(doc, lang):
    """Figure legends as a separate section after the tables (JGH)."""
    if lang != "en":
        return
    doc.add_heading("Figure Legends", level=1)
    legends = [
        (f"Figure 1. National dispensed quantity of hepatitis C antiviral drugs, indexed to "
         f"FY{Y0} = 100, with total interferon-free DAA dispensing on the right axis and "
         "dated approval and NHI-listing milestones. Ribavirin and conventional interferon "
         "are shown separately because they do not mark interferon-based hepatitis C therapy "
         "uniquely. Quantities are the published NDB totals; a series reaching zero means "
         "the drug fell below the NDB publication threshold in every table, not that "
         "dispensing ceased. DAA, direct-acting antiviral; FY, fiscal year; NDB, National "
         "Database of Health Insurance Claims and Specific Health Checkups of Japan; NHI, "
         "National Health Insurance."),
        (f"Figure 2. Interferon-free DAA dispensed quantity by product, fiscal years "
         f"{Y0}-{Y1}. Elbasvir and grazoprevir are shown separately because the NDB "
         "product files list the co-packaged components individually; ombitasvir/"
         "paritaprevir/ritonavir is a single co-formulated product. Products absent in a "
         "given year were below the NDB publication threshold in every table; the published "
         "totals are lower bounds. DAA, direct-acting "
         "antiviral; NDB, National Database of Health Insurance Claims and Specific Health "
         "Checkups of Japan."),
        ("Figure 3. Estimated national hepatitis C antiviral treatment volume "
         "(interferon-free DAA courses plus peginterferon courses) and the DAA share of it. "
         "Ribavirin is excluded because it accompanied both interferon-based and "
         "interferon-free regimens. DAA, direct-acting antiviral."),
    ]
    for lg in legends:
        p = doc.add_paragraph()
        p.add_run(lg)


def add_pds_titlepage(doc):
    """Pharmacoepidemiology & Drug Safety title-page front matter (English)."""
    meta = [
        ("Article type", "Original Report (observational, national open-data study)"),
        ("Running head", "Speed of the interferon-to-DAA switch in Japan"),
        ("Corresponding author", "[Name], [Affiliation], [Address], [Email]"),
        ("Authors / affiliations", "[To be completed by the authors]"),
    ]
    for k, v in meta:
        p = doc.add_paragraph()
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)


def add_keywords(doc):
    p = doc.add_paragraph()
    p.add_run("Keywords: ").bold = True
    p.add_run("hepatitis C; direct-acting antivirals; interferon; pharmacoepidemiology; "
              "prescription trends; drug utilization; NDB Open Data; Japan")


def add_key_points(doc):
    """P&DS 'Key Points' / take-home box."""
    doc.add_heading("Key Points", level=2)
    pts = [
        "National drug-utilization open data can show how quickly a population leaves an "
        "established standard therapy after a new option is approved and reimbursed.",
        "In Japan, peginterferon dispensing for hepatitis C fell "
        f"{fmt(peg_drop,1)}% between FY{Y0} and FY{Y1} and the first-generation protease "
        f"inhibitors given with it disappeared from FY{pi_zero_year}.",
        "On an estimated-course scale, interferon-free direct-acting antivirals accounted "
        f"for {fmt(share_knot,1)}% of hepatitis C antiviral treatment within one fiscal "
        f"year of the first interferon-free reimbursement, up from {fmt(share_first,1)}%.",
        "Estimated total treatment volume peaked in "
        f"FY{comb_peak_fy} and then fell {fmt(comb_fall,1)}%, the shape expected when a "
        "curative therapy works through a prevalent population.",
        "Findings describe national dispensed quantities, not patient counts; no measure "
        "of media coverage or of individual treatment decisions was available.",
    ]
    for t in pts:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(t)


def add_pls(doc):
    """Plain Language Summary (mandatory for P&DS; one paragraph, <=200 words).

    Placed immediately after Key Points. Numbers come from the results, rounded to
    lay-readable precision.
    """
    doc.add_heading("Plain Language Summary", level=2)
    p = doc.add_paragraph()
    p.add_run(
        "Hepatitis C is a long-lasting liver infection that can lead to cirrhosis and "
        "liver cancer. For years, treatment in Japan relied on interferon injections, "
        "which caused hard side effects and cured only some people. In 2014 and 2015, "
        "Japan approved tablet-only treatments that work without interferon and began "
        "paying for them through public insurance. We asked how quickly the country "
        "actually moved from the old treatment to the new one. Using national open "
        "records of medicines supplied across Japan from the "
        f"{Y0} to the {Y1} fiscal year, we followed the yearly amount of each drug and "
        "converted those amounts into approximate numbers of treatment courses. "
        f"Interferon injections almost vanished, falling {fmt(peg_drop,0)}% over the ten "
        "years, and the drugs given alongside them stopped being used altogether. Within "
        "a single year of insurance coverage, the new tablets made up about "
        f"{fmt(share_knot,0)}% of estimated treatment. Total treatment rose sharply, "
        f"peaked in {comb_peak_fy}, and then fell by about {fmt(comb_fall,0)}% \u2014 the "
        "pattern expected when a curative treatment works through a large group of people "
        "already living with an infection. These records count medicines rather than "
        "people, so they describe national trends, not individual decisions.")


def add_pds_statements(doc):
    """Ethics / consent / COI / funding / reporting-guideline statements."""
    items = [
        ("Ethics approval and consent",
         "This study used only publicly available, aggregated national open data "
         "(NDB Open Data) containing no individual-level or identifiable information; "
         "ethics-committee approval and informed consent were therefore not required."),
        ("Conflict of interest", "The authors declare no conflicts of interest."),
        ("Funding", "This research received no specific grant from any funding agency."),
        ("Reporting guideline",
         "This observational study is reported in line with the STROBE guideline for "
         "cross-sectional/ecological analyses of routinely collected aggregate data; a "
         "completed STROBE checklist is provided as Supplementary Material."),
    ]
    for h, t in items:
        p = doc.add_paragraph()
        p.add_run(f"{h}. ").bold = True
        p.add_run(t)


def _event_date(e):
    """Day-precise date from precision='day(YYYY-MM-DD)', else the month."""
    prec = str(e["precision"])
    if prec.startswith("day(") and prec.endswith(")"):
        return prec[4:-1]
    return str(e["event_month"])


def add_events_table(doc, lang):
    cols = (["FY", "Date", "Drug", "Milestone", "Source"] if lang == "en"
            else ["年度", "日付", "薬剤", "イベント", "出所"])
    t = doc.add_table(rows=1, cols=len(cols))
    t.style = "Table Grid"
    _set_no_vertical_borders(t)
    for j, c in enumerate(cols):
        r = t.rows[0].cells[j].paragraphs[0].add_run(c); r.bold = True
    drug_col = "drug_en" if lang == "en" else "drug_ja"
    src_col = "source_en" if lang == "en" else "source"
    mile = {"approval": "approval" if lang == "en" else "承認",
            "nhi_listing": "NHI listing" if lang == "en" else "薬価収載"}
    for _, e in EV.iterrows():
        cells = t.add_row().cells
        cells[0].text = str(int(e["fy"]))
        cells[1].text = _event_date(e)
        cells[2].text = str(e[drug_col])
        cells[3].text = mile.get(e["milestone"], e["milestone"])
        cells[4].text = str(e[src_col])
    for row in t.rows:
        for cell in row.cells:
            for pph in cell.paragraphs:
                for rr in pph.runs:
                    rr.font.size = Pt(7.5)


def add_quantity_table(doc, lang):
    """National dispensed quantity by drug group and fiscal year."""
    groups = ["IFN_peg", "PI_ifn", "IFN_conv", "ribavirin", "DAA"]
    head = (["Fiscal year", "Peginterferon (syringes)",
             "First-generation protease inhibitors (with peginterferon)",
             "Conventional interferon (not hepatitis-C-specific)",
             "Ribavirin (both regimens)", "Interferon-free DAA total"]
            if lang == "en"
            else ["年度", "ペグインターフェロン（本）",
                  "第一世代プロテアーゼ阻害薬（ペグIFN併用）",
                  "従来型インターフェロン（C型肝炎特異的でない）",
                  "リバビリン（旧・新両レジメン）", "IFNフリーDAA合計"])
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    _set_no_vertical_borders(t)
    for j, c in enumerate(head):
        rr = t.rows[0].cells[j].paragraphs[0].add_run(c); rr.bold = True
    for fy, row in TS.iterrows():
        cells = t.add_row().cells
        cells[0].text = str(int(fy))
        for j, g in enumerate(groups, 1):
            v = float(row[g])
            up = float(TS_UP.loc[fy, g])
            cells[j].text = f"NP (\u2264{fmt(up, 0)})" if v == 0 else fmt(v, 0)
    for row in t.rows:
        for cell in row.cells:
            for pph in cell.paragraphs:
                for rr in pph.runs:
                    rr.font.size = Pt(7.5)


def add_course_table(doc, lang):
    """Estimated treatment courses: DAA, peginterferon, combined volume, DAA share."""
    daa_c = COURSE["estimated_courses_by_fy"]
    daa_alt = COURSE["estimated_courses_by_fy_longer_duration"]
    peg_c = COURSE["peginterferon"]["estimated_courses_by_fy"]
    comb = CV["estimated_courses_by_fy"]
    share = CV["daa_share_by_fy"]
    head = (["Fiscal year", "Interferon-free DAA courses",
             "Interferon-free DAA courses (longer-duration assumption)",
             "Peginterferon courses", "Combined treatment volume", "DAA share (%)"]
            if lang == "en"
            else ["年度", "IFNフリーDAAコース数", "IFNフリーDAAコース数（長期投与仮定）",
                  "ペグインターフェロンコース数", "合算治療量", "DAAの割合（%）"])
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    _set_no_vertical_borders(t)
    for j, c in enumerate(head):
        rr = t.rows[0].cells[j].paragraphs[0].add_run(c); rr.bold = True
    for y in sorted(daa_c, key=int):
        cells = t.add_row().cells
        cells[0].text = str(y)
        cells[1].text = fmt(daa_c[y], 0)
        cells[2].text = fmt(daa_alt[y], 0)
        cells[3].text = fmt(peg_c[y], 0)
        cells[4].text = fmt(comb[y], 0)
        # share is undefined for a year with no estimated treatment at all
        cells[5].text = "-" if share[y] is None else fmt(share[y] * 100, 1)
    for row in t.rows:
        for cell in row.cells:
            for pph in cell.paragraphs:
                for rr in pph.runs:
                    rr.font.size = Pt(7.5)


def add_tables_to_doc(doc, lang, heading=True):
    """Add all three tables to a document (used for the tables-only docx and,
    for Hepatology Research, appended after the figure legends)."""
    if heading:
        doc.add_heading("Tables" if lang == "en" else "表", level=1)
    add_caption(doc, ("Table 1. Official approval and NHI drug-price listing milestones "
                      "for interferon-free direct-acting antivirals in Japan."
                      if lang == "en"
                      else "表1．日本におけるIFNフリー直接作用型抗ウイルス薬の承認・"
                           "薬価基準収載イベント。"))
    add_events_table(doc, lang)

    doc.add_page_break()
    add_caption(doc, ("Table 2. National dispensed quantity by drug group and fiscal year "
                      "(NDB Open Data). Units differ across groups and are not additive."
                      if lang == "en"
                      else "表2．薬剤群別・年度別の全国処方数量（NDBオープンデータ）。"
                           "群間で単位が異なり合算できない。"))
    add_quantity_table(doc, lang)

    doc.add_page_break()
    add_caption(doc, ("Table 3. Estimated treatment courses by fiscal year (estimates from "
                      "documented regimen durations, one anchor product per regimen; not "
                      "observed patient counts)." if lang == "en"
                      else "表3．年度別の推定治療コース数（文献記載の投与期間に基づく推定値、"
                           "レジメンごとに代表成分1つを計上。実測の患者数ではない）。"))
    add_course_table(doc, lang)


def build_tables_doc(lang):
    doc = Document()
    add_tables_to_doc(doc, lang, heading=True)
    path = os.path.join(OUT, f"tables_{lang}.docx")
    doc.save(path)
    print("wrote", path)


def build_pptx(lang):
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    figs = [
        (f"fig1_ifn_collapse_{lang}.png",
         "Figure 1. National dispensed quantity of hepatitis C antiviral drugs"
         if lang == "en" else "図1．C型肝炎抗ウイルス薬の全国処方数量"),
        (f"fig2_daa_wave_{lang}.png",
         "Figure 2. Interferon-free DAA dispensed quantity by product"
         if lang == "en" else "図2．IFNフリーDAAの製剤別処方数量"),
        (f"fig3_treatment_volume_{lang}.png",
         "Figure 3. Estimated treatment volume and the interferon-free DAA share of it"
         if lang == "en" else "図3．推定治療量とIFNフリーDAAの占める割合"),
    ]
    for fn, cap in figs:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(PInches(0.5), PInches(0.2), PInches(12.3), PInches(0.7))
        tf = tb.text_frame; tf.text = cap
        tf.paragraphs[0].runs[0].font.size = PPt(20)
        tf.paragraphs[0].runs[0].font.bold = True
        slide.shapes.add_picture(os.path.join(OUT, fn), PInches(1.4), PInches(1.1),
                                 height=PInches(5.7))
    path = os.path.join(OUT, f"figures_{lang}.pptx")
    prs.save(path)
    print("wrote", path)


def _apply_jepi_format(doc):
    """Journal of Epidemiology: simple single-column layout, 12 pt Times New
    Roman, double-spaced, 25 mm margins, page numbers. Continuous line
    numbering is applied only to the Abstract-through-Acknowledgments section,
    which is created as its own section by add_jepi_titlepage and closed before
    the references."""
    for section in doc.sections:
        section.top_margin = Mm(25)
        section.bottom_margin = Mm(25)
        section.left_margin = Mm(25)
        section.right_margin = Mm(25)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.paragraph_format.line_spacing = 2.0
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _add_page_number_header(doc)


def add_jepi_titlepage(doc, title):
    """Journal of Epidemiology title page, then a new section whose continuous
    line numbering covers Abstract through Acknowledgments (per the Guide for
    Authors, references and graphics carry no line numbers)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    p.paragraph_format.space_after = Pt(18)

    meta = [
        ("Article type", "Original Article"),
        ("Short title", "Interferon-to-DAA transition in Japan"),
        ("Authors", "[To be completed by the authors; given name(s) and family "
                    "name(s), with affiliation letters as superscripts; first "
                    "author ORCID iD linked in the submission system]"),
        ("Affiliations", "[To be completed by the authors; full postal address "
                         "including country for each affiliation]"),
        ("Corresponding author", "[Name], [Affiliation], [Full postal address "
                                 "including country], Tel: [telephone], "
                                 "Email: [email]"),
        ("Word count", "Abstract: {{ABSTRACT_WORDS}} words; main text: "
                       "{{MAIN_WORDS}} words (Introduction to Discussion)"),
        ("Numbers of tables, figures, and supplementary materials",
         "3 tables, 3 figures, 1 supplementary file (completed STROBE "
         "checklist)"),
    ]
    for k, v in meta:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)

    # Start the line-numbered section (Abstract -> Acknowledgments).
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    _set_line_numbering(sec)


def _add_jepi_abstract(doc, T):
    """JE structured abstract (Background / Methods / Results / Conclusions),
    <=250 words."""
    ab = doc.add_paragraph()
    ab.add_run(T["abs_bg"]).bold = True
    ab.add_run("How quickly a national population leaves an established "
               "treatment once a better option is approved and reimbursed is "
               "rarely measured from routine data. Using Japan's national "
               "claims open data, we described the countrywide transition from "
               "interferon-based to interferon-free hepatitis C therapy. ")
    ab.add_run(T["abs_me"]).bold = True
    ab.add_run(f"From the National Database of Health Insurance Claims and "
               f"Specific Health Checkups of Japan (NDB) Open Data (fiscal "
               f"years {Y0}-{Y1}), we extracted national dispensed quantities "
               "of peginterferon, first-generation protease inhibitors given "
               "with peginterferon, conventional interferon, ribavirin, and "
               f"{_num_word(n_daa)} interferon-free direct-acting antiviral "
               "products, respecting the database's publication thresholds. "
               "Quantities were converted to estimated treatment courses using "
               "documented regimen durations and described with segmented "
               "log-linear and exponential trend models. ")
    ab.add_run(T["abs_re"]).bold = True
    ab.add_run(f"Peginterferon dispensing fell {fmt(peg_drop,1)}% from fiscal "
               f"year {Y0} to {Y1} ({fmt(peg_r,0)}% per year, 95% confidence "
               f"interval {fmt(peg_lo,0)}-{fmt(peg_hi,0)}%). Estimated total "
               f"treatment volume rose from {fmt(comb_first,0)} to a fiscal "
               f"year {comb_peak_fy} peak of {fmt(comb_peak,0)} courses, then "
               f"fell {fmt(comb_fall,1)}% to {fmt(comb_last,0)} by fiscal year "
               f"{Y1}; the interferon-free share of estimated courses moved "
               f"from {fmt(share_first,1)}% to {fmt(share_knot,1)}% within one "
               "fiscal year. ")
    ab.add_run(T["abs_co"]).bold = True
    ab.add_run("National hepatitis C treatment switched almost entirely to "
               "interferon-free regimens within a year of reimbursement, and "
               "total volume surged then declined as the prevalent pool was "
               "treated. National open claims data can measure such "
               "transitions, provided their publication rules are handled "
               "explicitly.")


def add_jepi_keywords(doc):
    """JE keywords: 3-5 terms, avoiding 'and'/'of' constructions."""
    p = doc.add_paragraph()
    p.add_run("Keywords: ").bold = True
    p.add_run("hepatitis C; antiviral agents; interferons; drug utilization; "
              "Japan")
    p.paragraph_format.space_after = Pt(12)


def _add_jepi_introduction(doc, p, lang):
    """JE Introduction: national open data as an instrument of descriptive
    pharmacoepidemiology first, then the hepatitis C transition."""
    p.add_run("National claims databases have become a standard instrument of "
              "descriptive epidemiology for following how medical care is "
              "actually delivered")
    cite(p, ["ndb"], lang)
    p.add_run(", and Japan's National Database of Health Insurance Claims and "
              "Specific Health Checkups of Japan (NDB) Open Data release puts "
              "part of that instrument in the public domain: annual, "
              "nationwide dispensed quantities of prescription drugs by sex- "
              "and age-stratified table. Using it for epidemiology means "
              "accepting its publication rules as part of the measurement "
              "instrument—each table lists only the highest-ranked products of "
              "a therapeutic class and masks cells below a fixed threshold—so "
              "series built from it are bounded rather than exact, and the "
              "rules can be used to derive explicit bounds")
    cite(p, ["ndb"], lang)
    p.add_run(". Whether such data can capture a population-level treatment "
              "transition is a methodological question worth answering on a "
              "transition that actually happened.")

    p = doc.add_paragraph()
    p.add_run("Chronic hepatitis C provides such a transition. Until 2014, "
              "treatment in Japan was built around interferon, given for "
              f"{ifn_weeks_min}-{ifn_weeks_max} weeks with ribavirin and, from "
              "2011, with a first-generation non-structural protein 3/4A "
              "(NS3/4A) protease inhibitor; it was poorly tolerated and did "
              "not cure everyone")
    cite(p, ["jsh"], lang)
    p.add_run(". Interferon-free all-oral regimens then arrived in quick "
              "succession: daclatasvir plus asunaprevir, approved in July 2014, "
              "was the first regimen worldwide to dispense with both "
              "interferon and ribavirin")
    cite(p, ["bms", "kumada"], lang)
    p.add_run(", followed by sofosbuvir-based and pangenotypic regimens with "
              f"cure rates above {daa_cure_min_pct}% and treatment durations "
              f"of {daa_dur_min}-{daa_dur_max} weeks")
    cite(p, ["omata", "mizokami"], lang)
    p.add_run(". Japan combines universal health insurance with publicly "
              "funded hepatitis testing and a national subsidy program that "
              "caps out-of-pocket payments for antiviral therapy, and each "
              "direct-acting antiviral (DAA) was listed on National Health "
              "Insurance (NHI) within months of approval and brought into the "
              "subsidy program as it was listed")
    cite(p, ["tanaka_cm", "subsidy", "nhi"], lang)
    p.add_run(".")

    p = doc.add_paragraph()
    p.add_run("Once such a therapy is approved and reimbursed, two population-"
              "level trajectories are possible. Use of the prior standard may "
              "fall steeply—if it was being used for want of an alternative—"
              "and total treatment volume may surge as the diagnosed prevalent "
              "pool is treated, then decline toward the flow of new diagnoses; "
              "alternatively, practical friction in referral pathways, "
              "treatment capacity, or clinical caution would slow both. "
              "Distinguishing between them matters because the shape of the "
              "trajectory determines how much treatment capacity a program "
              "must mobilize and how soon its binding constraint shifts from "
              "treatment to case-finding. Japan has historically approved new "
              "drugs later than the United States or Europe")
    cite(p, ["tanaka_lag"], lang)
    p.add_run(", which makes the speed of this switch additionally "
              f"informative. We therefore used NDB Open Data (fiscal years "
              f"{Y0}-{Y1}) to describe how national dispensing of "
              "interferon-based therapy and of interferon-free DAAs changed "
              "around the approval and NHI-listing milestones of 2014-2017, "
              "how fast the two exchanged places on a common estimated-course "
              "scale, and how much total hepatitis C treatment volume the "
              "transition generated.")


def _add_jepi_discussion_opendata(doc, lang):
    """JE Discussion paragraph: what this study shows about the use of
    national open data for descriptive treatment epidemiology, including the
    censoring introduced by the publication rules."""
    p = doc.add_paragraph()
    p.add_run("The episode also illustrates what national open claims data can "
              "and cannot measure. The NDB Open Data release is complete in "
              "coverage—the underlying database captures essentially all "
              "insurance claims in Japan—but deliberately partial in "
              "publication: each table lists only the highest-ranked products "
              "of a class and shows small cells as a dash")
    cite(p, ["ndb"], lang)
    p.add_run(". The dominant features of this transition (the collapse of "
              "interferon-based therapy, the surge in DAA dispensing, and the "
              "subsequent decline) involve quantities far above the threshold, "
              f"so they survive those rules intact; the rules mattered only at "
              "the margins, where we treated unpublished cells as zero in the "
              "primary series and bounded them in the upper-bound sensitivity "
              f"series, which moved the post-peak decline by less than "
              f"{fmt(abs(comb_post_r - comb_post_r_up),1)} percentage points "
              "per year and changed no qualitative conclusion. The same "
              "construction—extracting complete products, flagging masked "
              "cells, and publishing bounds rather than point values—is "
              "applicable to any drug class tracked through the same data "
              "release, and makes the analysis reproducible end to end from "
              "public sources. The limits are equally instructive: dispensed "
              "quantity is not a patient count, annual resolution hides "
              "within-year timing around a listing date, and without "
              "individual-level linkage the data cannot show who started "
              "treatment, who deferred it, or why.")


def add_jepi_backmatter(doc, lang):
    """JE Acknowledgments (with funding, AI-use, and conflicts-of-interest
    statements) followed by the Data Availability section."""
    if lang != "en":
        return
    doc.add_heading("Acknowledgments", level=1)
    doc.add_paragraph(
        "Funding: This research did not receive any specific grant from "
        "funding agencies in the public, commercial, or not-for-profit "
        "sectors. During the preparation of this manuscript the authors used "
        "ChatGPT/GPT-4 for language editing and formatting assistance; after "
        "using this tool, the authors reviewed and edited the content and take "
        "full responsibility for the final manuscript.")
    p = doc.add_paragraph()
    p.add_run("Ethics approval. ").bold = True
    p.add_run("This study used only publicly available, aggregated national "
              "open data and involved no individual-level or identifiable "
              "information; ethics-committee approval was therefore not "
              "required.")
    p = doc.add_paragraph()
    p.add_run("Conflicts of interest. ").bold = True
    p.add_run("The authors declare they have no conflict of interest with "
              "respect to this research study and paper.")

    doc.add_heading("Data Availability", level=1)
    doc.add_paragraph(
        "Data derived from a source in the public domain: the raw NDB Open "
        "Data workbooks are published by the Ministry of Health, Labour and "
        "Welfare and can be re-downloaded with the provided script "
        "(scripts/download_ndb.py). All analysis code, derived datasets, and "
        "the figure and manuscript generators are openly available at "
        + REPO_URL + ", so the full pipeline (download, build, analyze, "
        "figures, manuscript) reproduces every reported number, figure, and "
        "table from the public data.")


def build_highlights():
    """JE Highlights: a separate editable file with 3-5 bullet points of at
    most 150 characters each (Guide for Authors requirement)."""
    bullets = [
        "National interferon-free DAA uptake replaced interferon-based "
        "hepatitis C therapy within one fiscal year.",
        "Estimated hepatitis C treatment volume surged, then declined as the "
        "diagnosed prevalent pool was treated.",
        "Japan's open claims data captured the transition despite "
        "top-ranked-product listing and small-cell masking rules.",
        "The trajectory shows how quickly reimbursed curative therapy can "
        "move a treated population.",
    ]
    for b in bullets:
        assert len(b) <= 150, f"highlight exceeds 150 chars ({len(b)}): {b}"
    doc = Document()
    doc.styles["Normal"].font.size = Pt(11)
    h = doc.add_paragraph()
    r = h.add_run("Highlights")
    r.bold = True
    r.font.size = Pt(12)
    for b in bullets:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(b)
    path = os.path.join(OUT, "highlights_jepi.docx")
    doc.save(path)
    print("wrote", path)


_JEPI_PAGE_TARGETS = {
    "Title page": ["Nationwide transition"],
    "Abstract": ["Abstract"],
    "Introduction": ["Introduction"],
    "Methods": ["Methods"],
    "Results": ["Results"],
    "Discussion": ["Discussion"],
    "Acknowledgments": ["Acknowledgments"],
    "Data Availability": ["Data Availability"],
    "References": ["References"],
    "Table 1": ["Table 1."],
    "Table 2": ["Table 2."],
    "Table 3": ["Table 3."],
    "Figure Legends": ["Figure Legends"],
}


def write_jepi_page_map():
    """Render the JE manuscript to PDF and record the page of each section
    heading, for the 'Reported on page No.' column of the STROBE checklist."""
    import subprocess
    src = os.path.join(OUT, "manuscript_en_jepi.docx")
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", OUT, src],
                   check=True, capture_output=True, timeout=300)
    pdf = os.path.join(OUT, "manuscript_en_jepi.pdf")
    info = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True,
                          check=True).stdout
    n_pages = int(re.search(r"Pages:\s+(\d+)", info).group(1))
    page_text = []
    for i in range(1, n_pages + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i),
                            "-layout", pdf, "-"],
                           capture_output=True, text=True, check=True).stdout
        page_text.append(t)
    pages = {}
    for key, pats in _JEPI_PAGE_TARGETS.items():
        found = None
        for i, text in enumerate(page_text, 1):
            # Strip the continuous line numbers the renderer adds (" 2   Abstract").
            lines = [re.sub(r"^\s*\d+\s+", "", ln).strip()
                     for ln in text.splitlines()]
            for pat in pats:
                if key == "Title page":
                    if any(pat in ln for ln in lines):
                        found = i
                        break
                elif key.startswith("Table "):
                    if any(ln.startswith(pat) for ln in lines):
                        found = i
                        break
                elif any(ln == pat for ln in lines):
                    found = i
                    break
            if found is not None:
                break
        pages[key] = found
    with open(JEPI_PAGEMAP_PATH, "w", encoding="utf-8") as fh:
        json.dump(pages, fh, indent=1)
    print("wrote", JEPI_PAGEMAP_PATH, pages)
    return pages


# STROBE item -> manuscript locations used for the JE 'Reported on page No.'
# column (multiple sections map to their page numbers, joined with commas).
_STROBE_JEPI_LOC = {
    "1": ["Title page", "Abstract"],
    "2": ["Introduction"],
    "3": ["Introduction"],
    "4": ["Methods"],
    "5": ["Methods", "Table 1"],
    "6": ["Methods"],
    "7": ["Methods"],
    "8": ["Methods", "Data Availability"],
    "9": ["Methods", "Discussion"],
    "10": ["Methods"],
    "11": ["Methods"],
    "12": ["Methods"],
    "13": ["Results"],
    "14": ["Results", "Table 2"],
    "15": ["Results", "Figure Legends"],
    "16": ["Results"],
    "17": ["Results"],
    "18": ["Discussion"],
    "19": ["Discussion"],
    "20": ["Discussion"],
    "21": ["Discussion"],
    "22": ["Acknowledgments"],
}


def _jepi_pages_str(item_no, pages):
    locs = _STROBE_JEPI_LOC.get(item_no, [])
    nums = sorted({pages[k] for k in locs if pages.get(k)})
    if not nums:
        return ""
    return "pp. " + ", ".join(str(n) for n in nums)


STROBE_ITEMS = [
    # (No., section, item text, our response / location)
    ("1", "Title and abstract",
     "(a) Indicate the study design with a commonly used term in the title or the "
     "abstract. (b) Provide in the abstract an informative and balanced summary of "
     "what was done and what was found.",
     "Title and structured abstract describe an observational, ecological (national "
     "aggregate) drug-utilization time-series analysis of NDB Open Data; the abstract "
     "summarizes design, data, methods and headline results."),
    ("2", "Background/rationale",
     "Explain the scientific background and rationale for the investigation being "
     "reported.",
     "Introduction explains why the speed at which a population leaves an established "
     "standard therapy after approval and reimbursement matters, and why the hepatitis C "
     "interferon -> interferon-free DAA transition is an informative setting."),
    ("3", "Objectives",
     "State specific objectives, including any prespecified hypotheses.",
     "Introduction states the two contrasting expectations: national dispensing of the "
     "prior standard therapy either falls steeply after the new option is approved and "
     "reimbursed, or changes slowly, indicating practical friction."),
    ("4", "Study design",
     "Present key elements of study design early in the paper.",
     "Methods: ecological/descriptive analysis of national annual dispensed quantity "
     f"(FY{Y0}-FY{Y1}); no individual-level data; no pre-intervention baseline within "
     "NDB, so descriptive trend models rather than a causal interrupted time series."),
    ("5", "Setting",
     "Describe the setting, locations, and relevant dates, including periods of "
     "recruitment, exposure, follow-up, and data collection.",
     "Methods: Japan, nationwide; NDB Open Data editions 1-10 mapped to fiscal years "
     f"{Y0}-{Y1}; approval/reimbursement milestones tabulated (Table 1)."),
    ("6", "Participants",
     "Give the eligibility criteria, and the sources and methods of selection of "
     "participants (cross-sectional).",
     "No individual participants: the unit is national aggregate dispensed quantity "
     "per drug per fiscal year, extracted for all records of the target products. "
     "Stated explicitly in Methods and Limitations (ecological data, not patient counts)."),
    ("7", "Variables",
     "Clearly define all outcomes, exposures, predictors, potential confounders, and "
     "effect modifiers.",
     "Methods: outcome = national dispensed quantity by drug group (peginterferon, "
     "conventional IFN, ribavirin, first-generation IFN-based protease inhibitors, "
     "interferon-free DAAs); milestones = dated approval and NHI drug-price listings; "
     "no measure of media coverage was available; potential secular confounders (e.g. "
     "COVID-19) and the hepatitis treatment subsidy program are discussed."),
    ("8", "Data sources/measurement",
     "For each variable of interest, give sources of data and details of methods of "
     "assessment (measurement).",
     "Methods and Data/code availability: MHLW NDB Open Data workbooks (sex/age x drug "
     "quantity tables); metric is total dispensed quantity (tablets/capsules or "
     "syringes/vials). Product classification listed; extraction code in the repository."),
    ("9", "Bias",
     "Describe any efforts to address potential sources of bias.",
     "Methods/Limitations: first-generation IFN-based protease inhibitors separated "
     "from interferon-free DAAs to avoid misclassification; units are not summed across "
     "products with different dosage units; absence of a pre-2014 baseline and lack of "
     "a control condition are stated as biases limiting causal inference."),
    ("10", "Study size",
     "Explain how the study size was arrived at.",
     "Methods: the study uses the complete set of national annual observations available "
     f"in NDB Open Data (n={n_obs} fiscal years); no sampling. Small n is flagged as a driver "
     "of wide uncertainty intervals."),
    ("11", "Quantitative variables",
     "Explain how quantitative variables were handled in the analyses.",
     "Methods: log-scale trend modelling; treatment-course sensitivity converts dispensed "
     "quantity to approximate courses using documented daily dose x duration "
     "(data/daa_course_assumptions.csv, data/ifn_course_assumptions.csv), counting one "
     "anchor product per two-drug regimen."),
    ("12", "Statistical methods",
     "(a) Describe all statistical methods, including those used to control for "
     "confounding. (b) sensitivity analyses.",
     "Methods: segmented (broken-stick) log-linear regression with a knot at the observed "
     "DAA peak, exponential/log-linear decay for IFN groups, Newey-West (HAC) standard "
     "errors and residual-bootstrap 95% intervals; duration-based sensitivity analysis "
     "for course estimates. Analysis is explicitly descriptive/non-causal."),
    ("13", "Participants (results)",
     "(a) Report numbers of individuals at each stage. (b) reasons for non-participation. "
     "(c) consider a flow diagram.",
     "Not applicable at the individual level (aggregate open data). Results and Methods "
     f"report the data units: {n_obs} fiscal years x drug groups; {n_daa} distinct "
     "interferon-free DAA products and 3 first-generation IFN-based protease inhibitors."),
    ("14", "Descriptive data",
     "(a) Give characteristics of study participants and information on exposures and "
     "potential confounders. (b) missing data. (c) follow-up (cohort).",
     "Results and Table 2: national dispensed quantity by group and fiscal year. Products "
     "absent in a given year appear as zero; no imputation. Milestones in Table 1."),
    ("15", "Outcome data",
     "Report numbers of outcome events or summary measures (cross-sectional).",
     "Results, Table 2, Figures 1-2: annual dispensed quantities and their changes; "
     "Figure 3 and Table 3 report estimated treatment courses and the DAA share of them "
     "(labelled estimates)."),
    ("16", "Main results",
     "(a) Give unadjusted and adjusted estimates and their precision. (b) category "
     "boundaries. (c) translate relative to absolute risk if relevant.",
     f"Results: peginterferon -{fmt(peg_drop,1)}% (FY{Y0} to FY{Y1}); protease inhibitors "
     f"below the NDB publication threshold from FY{pi_zero_year}; ribavirin (used in both "
     f"regimens) below the threshold from FY{rbv_zero_year}; DAA peak "
     f"FY{daa_peak_fy} then -{fmt(daa_fall,0)}% (upper-bound series -{fmt(daa_fall_up,0)}%); estimated combined treatment volume "
     f"peaked in FY{comb_peak_fy} and fell {fmt(comb_fall,1)}%; annual rates with "
     "HAC/bootstrap 95% intervals and slope-change P values from segmented regression."),
    ("17", "Other analyses",
     "Report other analyses done (e.g. subgroups, interactions, sensitivity analyses).",
     "Results: product-level DAA breakdown (Figure 2) and treatment-course sensitivity "
     "under baseline vs longer-duration assumptions (Table 3; cumulative "
     f"{fmt(course_total_lo,0)}-{fmt(course_total_hi,0)} estimated DAA courses)."),
    ("18", "Key results",
     "Summarise key results with reference to study objectives.",
     "Discussion opens by answering the question posed in the Introduction: dispensing "
     "shifted to interferon-free regimens within one fiscal year of reimbursement on the "
     "estimated-course scale, so annual-scale friction was small in this case."),
    ("19", "Limitations",
     "Discuss limitations, taking into account sources of potential bias or imprecision.",
     "Limitations: no pre-2014 baseline within NDB; dispensed quantity != patient counts; "
     "annual resolution precludes within-year ITS; small n; no control/placebo event; "
     "COVID-19 and other FY2020+ secular changes cannot be separated from the DAA decline; "
     "course figures are estimates dependent on regimen assumptions; ribavirin cannot be "
     "assigned to either regimen; media coverage and exposure to it were not measured."),
    ("20", "Interpretation",
     "Give a cautious overall interpretation considering objectives, limitations, "
     "multiplicity, and other relevant evidence.",
     "Discussion/Conclusions: interpretation limited to the speed of the observed "
     "displacement; the surge-then-decay pattern is presented as consistent with a "
     "curative therapy working through a prevalent pool and with accumulated demand, "
     "without claiming to distinguish them or any individual-level cause."),
    ("21", "Generalisability",
     "Discuss the generalisability (external validity) of the study results.",
     "Discussion: the speed observed here reflects a decisively better therapy, prompt "
     "NHI listing, and subsidized affordability in a population with a large diagnosed "
     "backlog; contrast with lower DAA uptake among insured US adults shows it may not "
     "generalize to other therapies or health systems."),
    ("22", "Funding",
     "Give the source of funding and the role of the funders for the present study and, "
     "if applicable, for the original study on which the present article is based.",
     "Funding statement: no specific grant. Data are public NDB Open Data (MHLW)."),
]


def build_strobe(journal=None):
    """Filled STROBE checklist (cross-sectional/ecological) as a supplementary docx.

    For journal='jepi' an extra 'Reported on page No.' column is added and
    filled from output/jepi_page_map.json (written by write_jepi_page_map()
    after rendering the manuscript to PDF).
    """
    pages = {}
    if journal == "jepi" and os.path.exists(JEPI_PAGEMAP_PATH):
        with open(JEPI_PAGEMAP_PATH, encoding="utf-8") as fh:
            pages = json.load(fh)

    doc = Document()
    doc.styles["Normal"].font.size = Pt(9.5)
    h = doc.add_paragraph()
    r = h.add_run("STROBE Statement — checklist of items for reports of "
                  "observational studies (cross-sectional / ecological adaptation)")
    r.bold = True; r.font.size = Pt(12)
    if journal == "jepi":
        doc.add_paragraph(
            "Strengthening the Reporting of Observational Studies in "
            "Epidemiology (STROBE). This study analyzes national aggregate "
            "drug-dispensing open data (no individual-level records); items "
            "are answered accordingly. Page numbers refer to the submitted "
            "main document (manuscript_en_jepi).")
    else:
        doc.add_paragraph(
            "Strengthening the Reporting of Observational Studies in Epidemiology (STROBE). "
            "This study analyzes national aggregate drug-dispensing open data (no "
            "individual-level records); items are answered accordingly, with locations given "
            "by manuscript section rather than page number.")
    ncols = 5 if journal == "jepi" else 4
    tbl = doc.add_table(rows=1, cols=ncols)
    tbl.style = "Table Grid"
    _set_no_vertical_borders(tbl)
    hdr = tbl.rows[0].cells
    headers = ["Item No.", "Section", "STROBE recommendation",
               "Location / response in manuscript"]
    if journal == "jepi":
        headers.append("Reported on page No.")
    for c, t in zip(hdr, headers):
        c.paragraphs[0].add_run(t).bold = True
    for no, sec, rec, resp in STROBE_ITEMS:
        cells = tbl.add_row().cells
        cells[0].text = no
        cells[1].text = sec
        cells[2].text = rec
        cells[3].text = resp
        if journal == "jepi":
            cells[4].text = _jepi_pages_str(no, pages)
    name = "strobe_checklist_jepi.docx" if journal == "jepi" else "strobe_checklist.docx"
    path = os.path.join(OUT, name)
    doc.save(path)
    print("wrote", path)


def build_cover_letter(journal="pds"):
    """Cover letter (English).

    Author/corresponding-author details are left as placeholders for local
    completion. Headline numbers are pulled from the results, not hard-coded.
    """
    doc = Document()
    doc.styles["Normal"].font.size = Pt(11)

    for line in ("[Author name]", "[Affiliation]", "[Address]", "[Email]", ""):
        doc.add_paragraph(line)
    doc.add_paragraph("[Date]")
    doc.add_paragraph()
    if journal == "hepres":
        recipient = ("The Editor-in-Chief", "Hepatology Research")
        title = ("Rapid population-level displacement of interferon-based hepatitis C "
                 "therapy by interferon-free direct-acting antivirals in Japan")
        article_type = "an Original Article"
        journal_name = "Hepatology Research"
    elif journal == "eid":
        recipient = ("The Editor-in-Chief", "Emerging Infectious Diseases")
        title = ("Rapid Population-Level Displacement of Interferon-Based Hepatitis C Therapy "
                 "by Interferon-Free Direct-Acting Antivirals, Japan")
        article_type = "a Research article"
        journal_name = "Emerging Infectious Diseases"
    elif journal == "jvh":
        recipient = ("The Editor-in-Chief", "Journal of Viral Hepatitis")
        title = (
            "Rapid Population-Level Displacement of Interferon-Based Hepatitis C Therapy "
            "by Interferon-Free Direct-Acting Antivirals in Japan"
        )
        article_type = "an Original Article"
        journal_name = "Journal of Viral Hepatitis"
    elif journal == "jgh":
        recipient = ("The Editors", "Journal of Gastroenterology and Hepatology")
        title = (
            "Rapid nationwide transition from interferon-based to interferon-free "
            "hepatitis C therapy in Japan: treatment dynamics after reimbursement and "
            "their relevance to Asia-Pacific elimination programs"
        )
        article_type = "an Original Article (Hepatology)"
        journal_name = "Journal of Gastroenterology and Hepatology"
    elif journal == "jepi":
        recipient = ("The Editor-in-Chief", "Journal of Epidemiology")
        title = (
            "Nationwide transition from interferon-based to interferon-free "
            "hepatitis C therapy in Japan: a descriptive study of national "
            "claims open data"
        )
        article_type = "an Original Article"
        journal_name = "Journal of Epidemiology"
    else:
        recipient = ("The Editor-in-Chief", "Pharmacoepidemiology and Drug Safety")
        title = TXT["en"]["title"]
        article_type = "an Original Report"
        journal_name = "Pharmacoepidemiology and Drug Safety"
    for line in recipient:
        doc.add_paragraph(line)
    doc.add_paragraph()

    doc.add_paragraph("Dear Editor,")

    p = doc.add_paragraph()
    p.add_run(
        f"Please consider our manuscript, \u201c{title},\u201d for "
        f"publication as {article_type} in {journal_name}.")

    # Cover letter: article category, word count, and reference count.
    if journal in ("eid", "jvh", "jgh"):
        counts_path = {"eid": EID_COUNTS_PATH, "jvh": JVH_COUNTS_PATH,
                       "jgh": JGH_COUNTS_PATH}[journal]
        if os.path.exists(counts_path):
            with open(counts_path, encoding="utf-8") as fh:
                counts = json.load(fh)
            wc_line = (
                f"The manuscript contains an abstract of {counts['abstract_words']} words, "
                f"a main text of {counts['main_words']} words, and {counts['reference_count']} references. "
                "All authors have approved the submission and declare no conflicts of interest."
            )
        else:
            wc_line = (
                "[Word count to be completed by the authors: abstract X words; main text Y words; "
                "references Z. All authors have approved the submission and declare no conflicts of interest.]"
            )
        p = doc.add_paragraph()
        p.add_run(wc_line)

    p = doc.add_paragraph()
    p.add_run(
        "Approval and reimbursement of a new therapy are recorded to the day, but how "
        "quickly an entire population actually moves away from the therapy it replaces is "
        "rarely measured. Hepatitis C offers an unusually clear setting: interferon-free "
        "direct-acting antivirals (DAAs) were approved and covered by National Health "
        "Insurance in Japan in 2014-2015 and took over from an interferon-based standard "
        "that had been in place for years.")

    p = doc.add_paragraph()
    p.add_run(
        f"Using Japan\u2019s NDB Open Data (fiscal years {Y0}\u2013{Y1}), we tracked "
        "national dispensed quantities of interferon-based therapy and of interferon-free "
        "DAAs against dated approval and reimbursement milestones, and converted "
        "quantities to estimated treatment courses so that a weekly injection and an "
        "8-week tablet regimen could be compared on one scale. Peginterferon dispensing "
        f"fell {fmt(peg_drop,1)}% and the protease inhibitors given with it disappeared "
        f"from FY{pi_zero_year}, while interferon-free regimens reached "
        f"{fmt(share_knot,1)}% of estimated treatment courses within one fiscal year of "
        "the first interferon-free listing. Estimated total treatment volume peaked in "
        f"FY{comb_peak_fy} and then fell {fmt(comb_fall,1)}%, the shape expected when a "
        "curative therapy works through a prevalent population. We report these as "
        "descriptive national trends with uncertainty intervals, discuss the subsidy "
        "program and other plausible contributors, and state explicitly that dispensed "
        "quantity is not a patient count and that no measure of media coverage or of "
        "individual treatment decisions entered the analysis.")

    p = doc.add_paragraph()
    if journal == "hepres":
        p.add_run(
            "We believe the work is suited to Hepatology Research. Japan's rapid transition "
            "from interferon-based to interferon-free hepatitis C therapy is a national "
            "example of the uptake problems faced by the global HCV elimination agenda, and "
            "the analysis connects clinical hepatology with health-system uptake by using "
            "open dispensing data to measure how quickly a population moved away from the "
            "established standard once the new regimens were listed and subsidized. The analysis is fully "
            "reproducible: all code and derived data are openly available at " + REPO_URL +
            ", and the raw NDB Open Data are re-downloadable from the Ministry of Health, "
            "Labour and Welfare, so every reported number, figure, and table can be "
            "regenerated from the public data. A completed STROBE checklist is provided as "
            "Supplementary Material.")
    elif journal == "eid":
        p.add_run(
            "We believe the work is suited to Emerging Infectious Diseases. Japan's rapid, "
            "nationwide displacement of interferon-based by interferon-free hepatitis C therapy "
            "illustrates how quickly a population can move to a curative regimen once it is "
            "approved, reimbursed, and affordable, a pattern relevant to the control and "
            "elimination of infectious diseases. The analysis is fully reproducible: all code "
            "and derived data are openly available at " + REPO_URL +
            ", and the raw NDB Open Data are re-downloadable from the Ministry of Health, "
            "Labour and Welfare, so every reported number, figure, and table can be "
            "regenerated from the public data. Figures are supplied as separate high-resolution "
            "files per the journal's Instructions for Authors. A completed STROBE checklist is "
            "provided as Supplementary Material.")
    elif journal == "jvh":
        p.add_run(
            "We believe the work is suited to Journal of Viral Hepatitis. Japan's rapid, "
            "nationwide displacement of interferon-based by interferon-free hepatitis C "
            "therapy offers a contemporary example of how quickly a national population can "
            "transition to a curative DAA regimen once it is approved and reimbursed, a "
            "question directly relevant to viral hepatitis treatment policy and elimination "
            "planning. The analysis is fully reproducible: all code and derived data are "
            "openly available at " + REPO_URL +
            ", and the raw NDB Open Data are re-downloadable from the Ministry of Health, "
            "Labour and Welfare, so every reported number, figure, and table can be "
            "regenerated from the public data. Figures are supplied as separate high-resolution "
            "files per the journal's Instructions for Authors. A completed STROBE checklist is "
            "provided as Supplementary Material.")
    elif journal == "jgh":
        p.add_run(
            "We believe the work is suited to the Journal of Gastroenterology and "
            "Hepatology and its Asia-Pacific readership. Japan is one of the few countries "
            "in the region on track for the 2030 elimination targets, and the manuscript "
            "documents, from national open dispensing data, what happened to treatment "
            "volume after interferon-free regimens were listed and subsidized: an almost "
            "complete switch within one fiscal year, a surge in total treatment, and a "
            "subsequent decline as the diagnosed pool was treated. We frame this as a "
            "planning template for elimination programs in the region, whose binding "
            "constraint shifts from treatment access to case-finding once such a switch "
            "is complete, while stating clearly what the aggregate data cannot show. The "
            "analysis is fully reproducible: all code and derived data are openly "
            "available at " + REPO_URL +
            ", and the raw NDB Open Data are re-downloadable from the Ministry of Health, "
            "Labour and Welfare, so every reported number, figure, and table can be "
            "regenerated from the public data. Figures are supplied as separate "
            "high-resolution files, tables follow the references, and a completed STROBE "
            "checklist is provided as Supplementary Material.")
    elif journal == "jepi":
        p.add_run(
            "We believe the work is suited to the Journal of Epidemiology. The "
            "manuscript is a descriptive epidemiological study of a national "
            "treatment transition built entirely from Japan's public claims "
            "open data, and it shows both what such data can measure—how fast a "
            "population leaves an established therapy once a better option is "
            "reimbursed, and how total volume evolves as the prevalent pool is "
            "treated—and how their publication rules (top-ranked listing and "
            "small-cell masking) should be handled. The analysis is fully "
            "reproducible: all code and derived data are openly available at "
            + REPO_URL +
            ", and the raw NDB Open Data are re-downloadable from the Ministry "
            "of Health, Labour and Welfare, so every reported number, figure, "
            "and table can be regenerated from the public data. Figures are "
            "supplied as separate files with their source files, tables are "
            "editable text, continuous line numbers run from the abstract "
            "through the acknowledgments, and a completed STROBE checklist "
            "with page numbers is provided as a supplementary file.")
    else:
        p.add_run(
            "We believe the work fits the scope of Pharmacoepidemiology and Drug Safety: it "
            "uses a national drug-utilization data source to quantify how quickly a population "
            "moves between therapeutic options around regulatory and reimbursement events, a "
            "question that recurs with every new therapeutic class. The analysis is fully "
            "reproducible: all code and derived data are openly available at " + REPO_URL +
            ", and the raw NDB Open Data are re-downloadable from the Ministry of Health, "
            "Labour and Welfare, so every reported number, figure, and table can be "
            "regenerated from the public data. A completed STROBE checklist is provided as "
            "Supplementary Material.")

    p = doc.add_paragraph()
    p.add_run(
        "This manuscript is original, has not been published previously, and is not under "
        "consideration elsewhere. All authors have approved the submission, contributed "
        "substantially to the study design, data interpretation, and manuscript preparation, "
        "and declare no conflicts of interest. The study used only publicly available, "
        "aggregated data with no individual-level information, so ethics-committee approval "
        "and informed consent were not required.")

    p = doc.add_paragraph()
    p.add_run("Author contributions. ").bold = True
    p.add_run(
        "[To be completed by the authors, per the ICMJE criteria. Example: T.O. conceived the "
        "study, designed the analysis, and drafted the manuscript; X.Y. curated the data and "
        "reviewed the analysis; Z.W. reviewed and edited the manuscript. All authors read "
        "and approved the final version.]")

    doc.add_paragraph("Thank you for considering our work.")
    doc.add_paragraph()
    doc.add_paragraph("Sincerely,")
    doc.add_paragraph("[Corresponding author name, on behalf of all authors]")
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Suggested reviewers. ").bold = True
    p.add_run("[Three potential referees with names, affiliations, and email addresses "
              "to be added by the authors, per the journal's Instructions for Authors.]")

    suffix = f"_{journal}" if journal else "_pds"
    path = os.path.join(OUT, f"cover_letter{suffix}.docx")
    doc.save(path)
    print("wrote", path)


if __name__ == "__main__":
    # `--only jgh` regenerates one journal's manuscript, cover letter, inline
    # review copy, and the shared tables/PPTX/STROBE files, leaving the other
    # journals' DOCX outputs untouched.
    _only = sys.argv[sys.argv.index("--only") + 1] if "--only" in sys.argv else None
    if _only:
        for lang in ("en", "ja"):
            build_tables_doc(lang)
            build_pptx(lang)
        if _only == "jepi":
            # The JE STROBE checklist needs real page numbers, so the
            # manuscript is rendered to PDF before the checklist is built.
            build_manuscript("en", journal="jepi")
            try:
                write_jepi_page_map()
            except Exception as exc:
                print("warning: could not render JE page map:", exc)
            build_strobe(journal="jepi")
            build_highlights()
            build_cover_letter(journal="jepi")
            build_manuscript("en", journal="jepi", inline=True)
        else:
            build_strobe()
            build_manuscript("en", journal=_only)
            build_cover_letter(journal=_only)
            build_manuscript("en", journal=_only, inline=True)
        raise SystemExit(0)
    for lang in ("en", "ja"):
        build_manuscript(lang)
        build_tables_doc(lang)
        build_pptx(lang)
    # Pharmacoepidemiology & Drug Safety submission variant (English):
    # structured abstract, Key Points, title-page front matter, STROBE/ethics/
    # COI/funding statements. Figures remain inline (P&DS accepts free-format
    # submission).
    build_manuscript("en", journal="pds")
    # Filled STROBE checklist (supplementary material for P&DS submission).
    build_strobe()
    # Cover letter for the P&DS submission (author details left as placeholders).
    build_cover_letter(journal="pds")

    # Hepatology Research submission variant (English):
    # structured abstract (Aim/Methods/Results/Conclusions), MeSH keywords,
    # abbreviations list, title page, acknowledgments/AI disclosure, and
    # figure legends after references.
    build_manuscript("en", journal="hepres")
    build_cover_letter(journal="hepres")

    # Inline review version of the Hepatology Research manuscript (not for
    # submission): same content with figures and tables embedded for readability.
    build_manuscript("en", journal="hepres", inline=True)

    # Emerging Infectious Diseases submission variant (English):
    # title page (article summary line, running title, MeSH keywords,
    # authors/affiliations/ORCID, corresponding author address, word counts),
    # unstructured 150-word abstract, IMRaD, acknowledgments, references,
    # tables, figure legends; figures supplied as separate files.
    build_manuscript("en", journal="eid")
    build_cover_letter(journal="eid")
    # Inline review version for EID: same content with figures and tables
    # embedded (not the submission file).
    build_manuscript("en", journal="eid", inline=True)

    # Journal of Viral Hepatitis submission variant (English):
    # title page, unstructured <=250-word abstract, up to 5 keywords,
    # abbreviations list, IMRaD, Statements and Declarations, AMA references,
    # tables and figure legends; figures supplied as separate files.
    build_manuscript("en", journal="jvh")
    build_cover_letter(journal="jvh")
    build_manuscript("en", journal="jvh", inline=True)

    # Journal of Gastroenterology and Hepatology submission variant (English,
    # Wiley transfer from JVH): title page with declarations, structured
    # <=250-word abstract (Background and Aim/Methods/Results/Conclusions),
    # 3-5 MeSH keywords, superscript Vancouver citations, 1.5 line spacing,
    # Introduction/Discussion reframed for Asia-Pacific HCV elimination,
    # tables and figure legends after references; figures as separate files.
    build_manuscript("en", journal="jgh")
    build_cover_letter(journal="jgh")
    build_manuscript("en", journal="jgh", inline=True)

    # Journal of Epidemiology submission variant (English): title page with
    # running title and display-item counts; structured abstract
    # (Background/Methods/Results/Conclusions, <=250 words); 3-5 keywords;
    # Introduction/Methods/Results/Discussion with the concluding paragraph
    # inside the Discussion; Acknowledgments (funding, AI use, conflicts of
    # interest) and Data Availability; continuous line numbers from the
    # abstract through the acknowledgments; AMA-style Vancouver superscript
    # references; tables (editable text) and figure legends after the
    # references; figures as separate files; mandatory Highlights file; STROBE
    # checklist with a 'Reported on page No.' column.
    build_manuscript("en", journal="jepi")
    try:
        write_jepi_page_map()
    except Exception as exc:
        print("warning: could not render JE page map:", exc)
    build_strobe(journal="jepi")
    build_highlights()
    build_cover_letter(journal="jepi")
    build_manuscript("en", journal="jepi", inline=True)
