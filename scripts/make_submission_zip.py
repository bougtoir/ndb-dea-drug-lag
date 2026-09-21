#!/usr/bin/env python3
"""
Assemble a journal submission package from the generated output/ files.

Supports:
  --journal pds      Pharmacoepidemiology & Drug Safety
  --journal hepres   Hepatology Research
  --journal eid      Emerging Infectious Diseases
  --journal jvh      Journal of Viral Hepatitis
  --journal jgh      Journal of Gastroenterology and Hepatology
  --journal jepi     Journal of Epidemiology

Run after make_figures.py and make_manuscript.py.
"""
import argparse
import os
import shutil
import zipfile

from PIL import Image

OUT = os.path.join(os.path.dirname(__file__), "..", "output")

JOURNALS = {
    "pds": {
        "pkg": "pds_submission",
        "files": [
            ("cover_letter_pds.docx", "00_cover_letter.docx"),
            ("manuscript_en_pds.docx", "01_manuscript_main.docx"),
            ("tables_en.docx", "02_tables.docx"),
            ("strobe_checklist.docx", "03_STROBE_checklist.docx"),
            ("figures_en.pptx", "04_figures_editable.pptx"),
        ],
        "figures": [
            ("fig1_ifn_collapse_en.png", "Figure1.png"),
            ("fig2_daa_wave_en.png", "Figure2.png"),
            ("fig3_treatment_volume_en.png", "Figure3.png"),
        ],
        "readme": """Pharmacoepidemiology and Drug Safety - submission package
========================================================

00_cover_letter.docx        Cover letter (author details to be completed locally)
01_manuscript_main.docx     Main manuscript: title page, structured abstract, Key
                            Points, Plain Language Summary, IMRaD text with figures
                            and tables inline, Word-native equations, references
02_tables.docx              Tables 1-3 as separate editable files
03_STROBE_checklist.docx    Completed STROBE checklist (Supplementary Material)
04_figures_editable.pptx    Figures 1-3, one per slide, editable
figures/Figure1-3.png       Individual figure files for upload

Every number, figure, and table is regenerated from the public NDB Open Data by the
pipeline in https://github.com/bougtoir/ndb-dea-drug-lag (download_ndb.py ->
build_dataset.py -> analyze.py -> course_estimate.py -> its_analysis.py ->
make_figures.py -> make_manuscript.py -> make_submission_zip.py).
""",
    },
    "hepres": {
        "pkg": "hepres_submission",
        "files": [
            ("cover_letter_hepres.docx", "00_cover_letter.docx"),
            ("manuscript_en_hepres.docx", "01_manuscript_main.docx"),
            ("tables_en.docx", "02_tables.docx"),
            ("strobe_checklist.docx", "03_STROBE_checklist.docx"),
            ("figures_en.pptx", "04_figures_editable.pptx"),
        ],
        "figures": [
            ("fig1_ifn_collapse_en.png", "Fig1.png"),
            ("fig2_daa_wave_en.png", "Fig2.png"),
            ("fig3_treatment_volume_en.png", "Fig3.png"),
        ],
        "readme": """Hepatology Research (Springer Nature) - submission package
=====================================

00_cover_letter.docx        Cover letter (author details to be completed locally)
01_manuscript_main.docx     Main manuscript: title page, structured abstract
                            (Aim/Methods/Results/Conclusions), MeSH keywords,
                            abbreviations list, IMRaD text, Statements and Declarations
                            (Funding, Competing Interests, Ethics approval, Consent,
                            Data and code availability, Authors' contributions),
                            references, and figure legends. Figures are supplied as
                            separate files per the Springer Nature guidelines.
02_tables.docx              Tables 1-3 as separate editable files
03_STROBE_checklist.docx    Completed STROBE checklist (Supplementary Material)
04_figures_editable.pptx    Figures 1-3, one per slide, editable
figures/Fig1-3.png          Individual figure files (high-resolution PNG)
figures/Fig1-3.jpg          Individual figure files (down-sampled JPG, <=1200 px width, for upload)

Every number, figure, and table is regenerated from the public NDB Open Data by the
pipeline in https://github.com/bougtoir/ndb-dea-drug-lag (download_ndb.py ->
build_dataset.py -> analyze.py -> course_estimate.py -> its_analysis.py ->
make_figures.py -> make_manuscript.py -> make_submission_zip.py).
""",
    },
    "eid": {
        "pkg": "eid_submission",
        "files": [
            ("cover_letter_eid.docx", "00_cover_letter.docx"),
            ("manuscript_en_eid.docx", "01_manuscript_main.docx"),
            ("tables_en.docx", "02_tables.docx"),
            ("strobe_checklist.docx", "03_STROBE_checklist.docx"),
            ("figures_en.pptx", "04_figures_editable.pptx"),
        ],
        "figures": [
            ("fig1_ifn_collapse_en.png", "Figure1.png"),
            ("fig2_daa_wave_en.png", "Figure2.png"),
            ("fig3_treatment_volume_en.png", "Figure3.png"),
        ],
        "high_res": True,
        "readme": """Emerging Infectious Diseases - submission package
=====================================

00_cover_letter.docx        Cover letter (article category, word and reference counts,
                            all-authors approval; author details to be completed locally)
01_manuscript_main.docx     Main manuscript: title page (article summary line, running
                            title, MeSH keywords, authors/affiliations/ORCID, corresponding
                            author address, word counts), unstructured 150-word abstract,
                            IMRaD text, acknowledgments, biographical sketch, address for
                            correspondence, references, tables (Word Table tool), and
                            figure legends. 12 pt Times New Roman, double-spaced, left
                            justified, continuous line numbering. Figures are NOT embedded
                            in the manuscript; they are supplied as separate files.
02_tables.docx              Tables 1-3 as a separate editable file
03_STROBE_checklist.docx    Completed STROBE checklist (Supplementary Material)
04_figures_editable.pptx    Figures 1-3, one per slide, editable
figures/Figure1-3.png       High-resolution individual figure files (PNG) for upload
figures/Figure1-3.jpg       High-resolution individual figure files (JPG, 300 dpi)

Every number, figure, and table is regenerated from the public NDB Open Data by the
pipeline in https://github.com/bougtoir/ndb-dea-drug-lag (download_ndb.py ->
build_dataset.py -> analyze.py -> course_estimate.py -> its_analysis.py ->
make_figures.py -> make_manuscript.py -> make_submission_zip.py).
""",
    },
    "jvh": {
        "pkg": "jvh_submission",
        "files": [
            ("cover_letter_jvh.docx", "00_cover_letter.docx"),
            ("manuscript_en_jvh.docx", "01_manuscript_main.docx"),
            ("manuscript_en_jvh_inline.docx", "01_manuscript_main_inline.docx"),
            ("tables_en.docx", "02_tables.docx"),
            ("strobe_checklist.docx", "03_STROBE_checklist.docx"),
            ("figures_en.pptx", "04_figures_editable.pptx"),
        ],
        "figures": [
            ("fig1_ifn_collapse_en.png", "Figure1.png"),
            ("fig2_daa_wave_en.png", "Figure2.png"),
            ("fig3_treatment_volume_en.png", "Figure3.png"),
        ],
        "tiffs": [
            ("fig1_ifn_collapse_en.tif", "Figure1.tif"),
            ("fig2_daa_wave_en.tif", "Figure2.tif"),
            ("fig3_treatment_volume_en.tif", "Figure3.tif"),
        ],
        "high_res": True,
        "readme": """Journal of Viral Hepatitis - submission package
=====================================

00_cover_letter.docx        Cover letter (word and reference counts, all-authors
                            approval; author details to be completed locally)
01_manuscript_main.docx     Main manuscript for submission: title page, keywords,
                            running title, authors/affiliations/ORCID, corresponding
                            author address, word counts, unstructured <=250-word
                            abstract, abbreviations list, IMRaD text, Statements and
                            Declarations, references, tables, and figure legends.
                            Figures are supplied as separate files per JVH guidelines.
01_manuscript_main_inline.docx For review only: same content with figures and tables
                            inline (not for submission in Editorial Manager).
02_tables.docx              Tables 1-3 as a separate editable file
03_STROBE_checklist.docx    Completed STROBE checklist (Supplementary Material)
04_figures_editable.pptx    Figures 1-3, one per slide, editable
figures/Figure1-3.png       High-resolution individual figure files (PNG) for upload
figures/Figure1-3.jpg       High-resolution individual figure files (JPG, 300 dpi)
figures/Figure1-3.tif       High-resolution individual figure files (TIFF, 300 dpi)

Every number, figure, and table is regenerated from the public NDB Open Data by the
pipeline in https://github.com/bougtoir/ndb-dea-drug-lag (download_ndb.py ->
build_dataset.py -> analyze.py -> course_estimate.py -> its_analysis.py ->
make_figures.py -> make_manuscript.py -> make_submission_zip.py).
""",
    },
    "jgh": {
        "pkg": "jgh_submission",
        "files": [
            ("cover_letter_jgh.docx", "00_cover_letter.docx"),
            ("manuscript_en_jgh.docx", "01_manuscript_main.docx"),
            ("manuscript_en_jgh_inline.docx", "01_manuscript_main_inline.docx"),
            ("tables_en.docx", "02_tables.docx"),
            ("strobe_checklist.docx", "03_STROBE_checklist.docx"),
            ("figures_en.pptx", "04_figures_editable.pptx"),
        ],
        "figures": [
            ("fig1_ifn_collapse_en.png", "Figure1.png"),
            ("fig2_daa_wave_en.png", "Figure2.png"),
            ("fig3_treatment_volume_en.png", "Figure3.png"),
        ],
        "tiffs": [
            ("fig1_ifn_collapse_en.tif", "Figure1.tif"),
            ("fig2_daa_wave_en.tif", "Figure2.tif"),
            ("fig3_treatment_volume_en.tif", "Figure3.tif"),
        ],
        "high_res": True,
        "readme": """Journal of Gastroenterology and Hepatology - submission package
===============================================================================

00_cover_letter.docx        Cover letter (word and reference counts, all-authors
                            approval; author details to be completed locally)
01_manuscript_main.docx     Main manuscript for submission: title page (article type,
                            short title, authors/affiliations/ORCID, corresponding
                            author, word counts, conflict of interest, financial
                            support, ethics, data availability), structured abstract
                            (Background and Aim / Methods / Results / Conclusions,
                            <=250 words), 3-5 MeSH keywords, Introduction / Methods /
                            Results / Discussion (<=3000 words, Conclusions inside the
                            Discussion), acknowledgments, Vancouver references with
                            superscript citation numbers, tables (titles above, no
                            vertical rules), and figure legends. 1.5 line spacing, US
                            spelling. Figures are supplied as separate files per JGH
                            guidelines.
01_manuscript_main_inline.docx For review only: same content with figures and tables
                            inline after their first citation (not for upload).
02_tables.docx              Tables 1-3 as a separate editable file
03_STROBE_checklist.docx    Completed STROBE checklist (Supplementary Material)
04_figures_editable.pptx    Figures 1-3, one per slide, editable (English)
figures/Figure1-3.png       High-resolution individual figure files (PNG) for upload
figures/Figure1-3.jpg       High-resolution individual figure files (JPG, 300 dpi)
figures/Figure1-3.tif       High-resolution individual figure files (TIFF, 300 dpi)

Every number, figure, and table is regenerated from the public NDB Open Data by the
pipeline in https://github.com/bougtoir/ndb-dea-drug-lag (download_ndb.py ->
build_dataset.py -> analyze.py -> course_estimate.py -> its_analysis.py ->
make_figures.py -> make_manuscript.py -> make_submission_zip.py --journal jgh).
Literature-derived numbers used in the Introduction/Discussion are kept in
data/discussion_support.json with their source keys.
""",
    },
    "jepi": {
        "pkg": "jepi_submission",
        "files": [
            ("highlights_jepi.docx", "01_highlights.docx"),
            ("manuscript_en_jepi.docx", "02_manuscript_main.docx"),
            ("manuscript_en_jepi_inline.docx", "02_manuscript_main_inline.docx"),
            ("cover_letter_jepi.docx", "03_cover_letter.docx"),
            ("tables_en.docx", "04_tables.docx"),
            ("strobe_checklist_jepi.docx", "05_STROBE_checklist.docx"),
            ("figures_en.pptx", "06_figures_source.pptx"),
        ],
        "figures": [
            ("fig1_ifn_collapse_en.png", "Figure1.png"),
            ("fig2_daa_wave_en.png", "Figure2.png"),
            ("fig3_treatment_volume_en.png", "Figure3.png"),
        ],
        "tiffs": [
            ("fig1_ifn_collapse_en.tif", "Figure1.tif"),
            ("fig2_daa_wave_en.tif", "Figure2.tif"),
            ("fig3_treatment_volume_en.tif", "Figure3.tif"),
        ],
        "high_res": True,
        "readme": """Journal of Epidemiology (Japan Epidemiological Association) - submission package
================================================================================

Upload order per the Guide for Authors (ScholarOne): Highlights, main document,
tables/figures, supplementary files.

01_highlights.docx          Highlights (mandatory): 4 bullet points, each <=150
                            characters
02_manuscript_main.docx     Main manuscript: title page (article type, running
                            title <=8 words, authors/affiliations, corresponding
                            author, word counts, numbers of tables/figures/
                            supplementary materials), structured abstract
                            (Background/Methods/Results/Conclusions, <=250
                            words), 5 keywords, Introduction / Methods / Results /
                            Discussion (main text <=3,500 words, concluding
                            statement inside the Discussion), Acknowledgments
                            (funding, AI-use, and conflicts-of-interest
                            statements), Data Availability, AMA-style Vancouver
                            references with superscript citation numbers,
                            tables (editable text, no vertical rules), and
                            figure legends. Continuous line numbers run from
                            the Abstract through the Acknowledgments. Figures
                            are NOT embedded; they are supplied as separate
                            files per the Guide for Authors.
02_manuscript_main_inline.docx For review only: same content with figures and
                            tables inline (not for upload).
03_cover_letter.docx        Cover letter (journal name, title, principal
                            findings and their significance for epidemiology,
                            all-authors approval; author details to be
                            completed locally)
04_tables.docx              Tables 1-3 as a separate editable file (also at the
                            end of the main manuscript)
05_STROBE_checklist.docx    Completed STROBE checklist with a 'Reported on page
                            No.' column (supplementary file)
06_figures_source.pptx      Figures 1-3 source file, one per slide, editable
                            (English) - submitted as the figures' original
                            (source) files
figures/Figure1-3.png       Individual figure files (PNG)
figures/Figure1-3.jpg       Individual figure files (JPG, 300 dpi)
figures/Figure1-3.tif       Individual figure files (TIFF, 300 dpi)

Every number, figure, and table is regenerated from the public NDB Open Data by
the pipeline in https://github.com/bougtoir/ndb-dea-drug-lag (download_ndb.py ->
build_dataset.py -> analyze.py -> course_estimate.py -> its_analysis.py ->
make_figures.py -> make_manuscript.py -> make_submission_zip.py --journal jepi).
Literature-derived numbers used in the Introduction/Discussion are kept in
data/discussion_support.json with their source keys.
""",
    },
}


def _convert_figures(fig_list, figure_dir, high_res=False):
    """Copy PNG originals and create JPG copies for journal upload.

    For most journals the JPG is down-sampled to <=1200 px at 96 dpi.
    For EID, which requests high-resolution (>=300 dpi at 5 inches or more)
    separate figure files, the original PNG size is retained and tagged
    as 300 dpi in the JPG copy.
    """
    for src, dst in fig_list:
        src_path = os.path.join(OUT, src)
        png_path = os.path.join(figure_dir, dst)
        shutil.copy(src_path, png_path)
        jpg_name = os.path.splitext(dst)[0] + ".jpg"
        jpg_path = os.path.join(figure_dir, jpg_name)
        with Image.open(src_path) as im:
            if im.mode == "RGBA":
                im = im.convert("RGB")
            if high_res:
                # Keep original pixel dimensions; EID figures are ~2000 px wide,
                # which is 5+ inches at 300 dpi.
                im.save(jpg_path, "JPEG", quality=95, dpi=(300, 300))
            else:
                max_width = 1200
                w, h = im.size
                if w > max_width:
                    ratio = max_width / w
                    im = im.resize((max_width, int(h * ratio)), Image.LANCZOS)
                im.save(jpg_path, "JPEG", quality=90, dpi=(96, 96))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--journal", choices=list(JOURNALS), default="hepres")
    args = parser.parse_args()

    cfg = JOURNALS[args.journal]
    pkg = os.path.join(OUT, cfg["pkg"])
    if os.path.isdir(pkg):
        shutil.rmtree(pkg)
    fig_dir = os.path.join(pkg, "figures")
    os.makedirs(fig_dir)

    with open(os.path.join(pkg, "00_README_submission.txt"), "w",
              encoding="utf-8") as fh:
        fh.write(cfg["readme"])

    for src, dst in cfg["files"]:
        shutil.copy(os.path.join(OUT, src), os.path.join(pkg, dst))

    _convert_figures(cfg["figures"], fig_dir, high_res=cfg.get("high_res", False))
    for src, dst in cfg.get("tiffs", []):
        shutil.copy(os.path.join(OUT, src), os.path.join(fig_dir, dst))

    zpath = os.path.join(OUT, f"{cfg['pkg']}.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _dirs, files in os.walk(pkg):
            for f in sorted(files):
                full = os.path.join(root, f)
                z.write(full, os.path.relpath(full, OUT))
    print("wrote", zpath)


if __name__ == "__main__":
    main()
