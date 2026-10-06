# PRE_REVISION_BASELINE — Hepatology International package

Audit date: 2026-09-25 (before revision). Branch: devin/1790332363-hi-submission.

## Canonical files
- Main manuscript: `output/manuscript_en_hi.docx` (inline copy `manuscript_en_hi_inline.docx`)
- Supplement: none (no separate supplement file in this package)
- Code: `scripts/{analyze,course_estimate,its_analysis,make_figures,make_manuscript,make_submission_zip}.py`
- Data: `data/*.csv` (hcv_timeseries, hcv_product_timeseries, +upper bounds, course assumptions, thresholds, events)
- Results: `results/{summary,course_estimate,its_summary}.json`
- Figures: `output/fig{1,2,3}_*_en.{png,pdf,tif}` + `figures_en.pptx`
- Tables: `output/tables_en.docx`
- Graphical abstract: `output/graphical_abstract_en.{png,jpg,tif}` (embedded in manuscript + separate file)
- Cover letter: `output/cover_letter_hi.docx`
- Checklist: `output/strobe_checklist_hi.docx`
- ZIP: `output/hi_submission.zip`

## Current title
"Rapid nationwide transition from interferon-based to interferon-free hepatitis C therapy in Japan: treatment dynamics after reimbursement and their relevance to Asia-Pacific elimination programs"

## Current abstract (230 words; Background and Aims / Methods / Results / Conclusions)
Describes nationwide transition IFN→IFN-free; FY2014–2023 NDB Open Data; pegIFN −99.2% (−36%/yr, CI 24–47%); PIs below threshold from FY2016; total volume 26,307 → peak 92,589 (FY2015) → −88.9% to 10,275 (FY2023); IFN-free share 50.9%→97.0% within 1 FY, stays >97%.

## Study period: FY2014–2023 (NDB editions 1–10), n=10 annual observations

## Primary numerical claims (from results JSONs)
- PegIFN dispensing −99.2% FY2014→2023; annual change −36.4% (HAC 95% −47.2..−23.5; boot −44.7..−27.1), p=1.5e-06
- PI fell below publication threshold from FY2016
- Estimated DAA courses: FY2014 13,404 → peak FY2015 89,858 → FY2023 10,174
- Combined estimated volume: 26,307 (2014) → 92,589 peak (2015) → 10,275 (2023); −88.9%
- IFN-free share of combined courses: 50.9% (2014) → 97.0% (2015), stays >97%
- Segmented combined volume: knot=FY2015 (observed peak, not estimated), pre-peak n=2 obs (+182%/yr, "two-point increment" caution already in its_summary), post-peak −24.7%/yr (boot −27.5..−21.8); slope_change p=1.55e-21 ← PHASE-3 TARGET
- Upper-bound masking sensitivity: post-peak −21.4%/yr; share changes <0.1pp
- Cumulative ≈297,637 estimated courses FY2014–2023 (range 265,914–297,637)

## Figure/table inventory
- Fig 1: indexed dispensed quantity IFN markers + DAA total (right axis) + milestones
- Fig 2: DAA dispensed quantity by product
- Fig 3: combined estimated courses + DAA share
- Tables 1–3: milestones; dispensed quantity by group; estimated courses
- GA: single-panel 3-cell schematic

## Reference list: 22 refs (NLM, numbered [n]); key = NDB [17], guidelines [10], Suoh NOT cited yet; Yamashita 2025 [21] cited

## Format status vs HI rules
Compliant: abstract ≤250; main+refs 3,804 ≤4,000; 22 ≤30 refs; 6 display items + GA embedded + separate file; Declarations before References; title-page ethics/COI/consent; bracket citations. STROBE generic 4-col.

## Known defects flagged by requester
- Equation rendering defect around hinge term (t−t0)+ and its definition (paras 27–29; inline OMML not extracting as text)
- GA: stray black horizontal line upper-right panel; 2014/2015 x-label crowding
- Slope-change P values (p<0.001 claims) must be de-emphasized (Phase 3)
- Suoh et al. 2026 not yet cited; intro must not claim "first nationwide DAA study"
