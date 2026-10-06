# FINAL HANDOFF — PDS submission variant (ndb-dea-drug-lag)

Built 2026-10-06 on branch `devin/1791265088-pds-submission` of `bougtoir/wip`,
project `ndb-dea-drug-lag`. One VM, one working tree; no new analysis; the
existing FY2014–FY2024 frozen results and reporting are reused verbatim.

## What changed (reframing only)

- Title → therapeutic-substitution framing.
- Structured abstract rewritten Purpose/Methods/Results/Conclusions (234 words).
- Introduction rewritten (5 paragraphs: uptake ≠ replacement → common scale
  (WHO DDD logic) → the IFN→DAA case → prior national work differentiated
  (Yamashita 2025, Suoh 2026) → objective).
- Discussion rewritten (6 paragraphs: replacement + post-substitution
  throughput decomposition; course-equivalent scale; prior-work
  differentiation; prevalent-pool interpretation; contributing factors;
  transferability).
- Conclusions rewritten for PDS; Methods gains two sentences (WHO
  course-equivalent logic; Wagner/Bernal segmented-regression citations).
- Title page, Key Points, statements (transparency/reproducibility added),
  cover letter (single-author "I" voice, ~1.5 pages, PDS-fit argument,
  Suoh/Yamaha differentiation, no mention of prior submissions) rewritten.
- `_apply_pds_format`: Times New Roman 12 pt, 1.5 spacing, 25 mm margins,
  page numbers (PDS free-format).
- `make_submission_zip.py --journal pds` → `PDS_submission_FINAL.zip` with
  FINAL-named contents; markdown docs ride along via `optional` file entries.

## Word/reference counts (machine-generated `output/pds_word_counts.json`)

- Abstract: 234 words (≤250). PLS: 190 words (≤200). Keywords: 7.
- Main text: 3,121 including 135 words of declaration statements →
  ≈2,986 body words vs the ~3,000 PDS Original Article guideline.
- References: 18, Vancouver order-of-appearance, superscript in-text.

## Numbers preserved (verified against structured outputs)

All statistics come from `results/summary.json`, `results/its_summary.json`,
`results/course_estimate.json`, `data/hcv_timeseries.csv`, and
`data/discussion_support.json` at build time — nothing hard-coded:
peginterferon −99.3% (34%/yr, 95% CI 22–44%, P<0.001); courses 26,307 →
92,589 (FY2015) → 7,950 (−91.4%, −23.4%/yr, CI 20.7–26.0%, P<0.001); DAA
share 50.9%→97.0% within one FY; cumulative 270,939–305,499 courses;
slope-change −1.26 log units/yr (CI −1.55 to −0.97, P<0.001) reported with
model-based calibration; n_obs=11. P values retained everywhere.

## Technical QA performed

- OMML equations present (32 `m:oMath`); renders correctly in Word/LibreOffice
  PDF export (17 pages).
- No stray `**`, `¿`, or placeholder text outside `[To be completed]` markers.
- Citations strictly in first-appearance order; Figure 1–3 and Tables 1–3 all
  cited in text; section order Abstract→Key Points→PLS→Introduction→Methods→
  Results→Discussion→Limitations→Conclusions→Statements→References.
- Suoh et al. 2026 differentiated in Introduction and Discussion (fair,
  non-adversarial: "a different question"; "complement rather than duplicate").
- No HI/JEPI/JGH/JMAJ/HEPRES/EID/JVH output file was overwritten (`--only pds`).

## Rebuild commands

    .venv/bin/python scripts/make_manuscript.py --only pds
    .venv/bin/python scripts/make_submission_zip.py --journal pds
    # full pipeline also builds every other journal variant unchanged:
    .venv/bin/python scripts/make_manuscript.py

## Unresolved items requiring human input

1. Author block: `[Name]`, `[Affiliation]`, `[Address]`, `[Email]`, `[ORCID]`
   placeholders on the title page and cover letter.
2. Author-contributions paragraph in the cover letter (ICMJE template).
3. Suggested reviewers (≤3) — placeholder in cover letter.
4. PDS-specific conflict-of-interest form — download/sign in Research Exchange.
5. If a graphical abstract is desired (optional at PDS), none is supplied.
