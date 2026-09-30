# FINAL HANDOFF — Hepatology International submission package

Date: 2026-09-21. Branch: `devin/1790332363-hi-submission` (updates PR #542, wip repo `ndb-dea-drug-lag/`).

## Package contents (`output/HI_submission_FINAL.zip`)

| File | Content |
|---|---|
| HI_MAIN_MANUSCRIPT_FINAL.docx | Title page, structured abstract (230 w), embedded Graphical Abstract, 10 keywords, IMRaD (Materials and Methods), Declarations, 23 NLM refs, Tables 1–3, figure legends |
| HI_MANUSCRIPT_INLINE_review_only.docx | Same manuscript with figures/tables inline — review convenience copy, not for upload |
| HI_COVER_LETTER_FINAL.docx | Cover letter (substitution dynamics / course-equivalent framing, explicit Suoh differentiation, Asia-Pacific relevance; no submission history) |
| HI_SUPPLEMENT_FINAL.docx | Completed STROBE checklist |
| HI_TABLES.docx | Tables 1–3, editable text |
| HI_FIGURES_SOURCE.pptx | Figures 1–3, editable, English |
| HI_GRAPHICAL_ABSTRACT_FINAL.tif / .png | Mandatory single-panel GA, 300 dpi |
| figures/Figure1–3 .tif/.png | Individual figure files (300 dpi TIFF) |

## Key facts

- **Period:** FY2014–FY2024 — all 11 published NDB Open Data editions (1st–11th). FY2024 uses the official MHLW 11th edition; its listing expanded to ALL products per class (no top-N cap), so the small ribavirin reappearance (28,508 units) is documented as a visibility change, not resumed use. See `FY2024_UPDATE_AUDIT.md`.
- **Novelty framing:** the dynamic replacement of an incumbent treatment standard (peginterferon-based) by a new standard (interferon-free DAAs) on a common course-equivalent scale — complementary to Suoh et al. 2026 (utilization/cost/geography), differentiated explicitly in Intro, Discussion, and cover letter. See `NEAR_NEIGHBOR_COMPARISON.md`.
- **Headline results (all generated from `results/`, none hard-coded):** peginterferon −99.3% (FY2014→FY2024); combined estimated courses 26,307 → FY2015 peak 92,589 → FY2024 7,950 (−91.4%); IFN-free share 50.9% → 97.0% within one year, ≥97% thereafter; post-peak decline ≈23.4%/yr; cumulative DAA volume 270,939–305,499 course-equivalents; n=11 annual observations.
- **P-value status (final finishing pass):** P values were retained/restored, not removed. Effect estimates are reported with 95% CIs and model-based P values where the fitted model provides them — exponential-decay rates (peginterferon −34%/yr, conventional IFN −25%/yr, ribavirin −50%/yr; all P<0.001), segmented-model post-peak slopes (DAA −23.8%/yr, combined −23.4%/yr; both P<0.001), and the slope-change terms (DAA −0.77 log units/yr, 95% CI −1.21 to −0.32, P<0.001; combined −1.26, 95% CI −1.55 to −0.97, P<0.001). Slope-change inference is explicitly model-based/exploratory — the knot was fixed at the observed FY2015 peak and the pre-peak segment has only 2 observations — stated in Methods, Results, Limitations, and STROBE item 12; the P values are not presented as formal change-point evidence. No new model family was added; the post-peak slope P is the β1+β2 HAC contrast of the same fitted segmented model.
- **Equation repair:** hinge notation `(t−t0)+ = max(0,t−t0)` rebuilt as native OMML; the superscript "+" is forced upright (`m:nor` + `m:sty="p"`), which eliminates the `¿¿` placeholder artifacts in LibreOffice/Word PDF export. Verified: PDF render contains zero `¿` characters.
- **Graphical Abstract:** stray arrow/line removed; axis ticks fixed to avoid 2014/2015 crowding; FY2024 included; single panel, 300 dpi, placed immediately after the abstract and labeled "Graphical Abstract".
- **Word/table/figure counts:** abstract 234 ≤250; main text 3,332, +refs 3,993 ≤4,000; 23 refs ≤30; 3 tables + 3 figures = 6 ≤6 combined; 10 keywords. Declarations before References. Vancouver numbering verified sequential in order of first appearance.
- **Final wording pass:** only three claim-calibration wording edits were made after the P-value pass — Introduction "binding constraint shifts" → "principal programmatic constraint may shift from treatment delivery toward case-finding"; Discussion "how soon does the binding constraint move" → "how soon might the principal programmatic constraint move ... toward diagnosis and linkage to care?"; Conclusion/cover letter "as expected when a curative therapy works through a prevalent population" → "a pattern consistent with progressive treatment of a prevalent population". No analysis was rerun; all effect estimates, 95% CIs, and P values preserved; figures/tables/equations unchanged; clean build completed (main+refs 3,993 words; PDF render has zero `¿`).
- **Cross-checks:** `REFERENCE_UPDATE_AUDIT.md` (ref list + Suoh addition), `HOSTILE_REVIEW_FINAL.md` (10-attack simulated referee pass, all answered or calibrated), `PRE_REVISION_BASELINE.md`.
- **Unresolved items:** author names/affiliations, corresponding-author block, author-contribution text, and suggested reviewers are `[...]` placeholders in the cover letter / title page for local completion. Declarations state the standard no-funding / no-COI / no-ethics-required text — confirm before submitting. JMA Journal variant (`jmaj`) was researched as a fallback (fully free OA, JE-like format) but not built.

## Reproduce

```
python3 scripts/download_ndb.py        # raw workbooks -> data/ndb_raw/ (editions 1-11)
python3 scripts/build_dataset.py
python3 scripts/analyze.py
python3 scripts/course_estimate.py
python3 scripts/its_analysis.py
python3 scripts/make_figures.py --lang en   # (+ --lang ja)
python3 scripts/make_manuscript.py --only hi
python3 scripts/make_submission_zip.py --journal hi
```

Submission route: Editorial Manager, subscription track (hybrid journal — no APC unless Open Choice OA is selected).
