# Critical reviewer review — EID submission

Journal: *Emerging Infectious Diseases* (Research article, 3,500-word text, 150-word abstract, ≤50 refs).
Manuscript: "Rapid Population-Level Displacement of Interferon-Based Hepatitis C Therapy by Interferon-Free Direct-Acting Antivirals, Japan."

## 1. Manuscript (novelty, focus, logic)

*Strengths*
- Clear ecological question: how fast national dispensing switched after reimbursement.
- Directly links approval/reimbursement dates to aggregate dispensing.
- Conclusions are scaled to the evidence: descriptive, no causal claim.

*Reviewer concerns / actions*
- **(High)** The study is ecological and offers no individual-level or mechanistic data. The framing repeatedly states what "cannot be separated" (media, COVID-19, secular trends), which is appropriate but may read as overly defensive. Consider keeping only the most consequential limitations.
- **(Medium)** The title says "Rapid Population-Level Displacement"; this is supported by the DAA share shift, but the data are dispensed *quantities*, not patients. The abstract and conclusion already caveat this. Verify that every summary sentence retains the "estimated treatment course" qualifier.
- **(Medium)** The Discussion compares the Japanese pattern with the US 23-35% one-year treatment rate. This is post-hoc and across healthcare systems; the paragraph should not overstate it as a contrast in effectiveness or policy. Current wording is cautious, but a reviewer may ask to separate this as an illustrative comparison rather than evidence.
- **(Fixed in this build)** DAA was not defined at first use in the main text; now expanded to "interferon-free direct-acting antivirals (DAAs)" in the Introduction.

## 2. Statistical design and analysis

*Strengths*
- Explicitly notes the absence of a pre-DAA baseline inside NDB.
- Uses both Newey-West and bootstrap intervals.
- Reports slope-change p-values and cautions against interpreting the pre-peak segment.

*Reviewer concerns / actions*
- **(High)** n = 10 annual observations. Any model with a knot is under-identified. The pre-peak segment uses only two observations and is correctly flagged as not a rate. The manuscript must not present the 71% pre-peak DAA increase as a reliable trend; current wording is "not interpreted as a rate." Keep this language.
- **(Medium)** Exponential decay fits for interferon-based drugs are a reasonable descriptive summary, but the model assumes a constant proportional decline. A reviewer could ask whether a segmented or non-parametric fit was considered. The current text can add a brief sentence that the model was chosen for parsimony given n=10 and that visual inspection supports monotonic decline.
- **(Low)** Confidence intervals are reported as 95% CI; source says they are Newey-West HAC or bootstrap. Make sure the Methods sentence explicitly states this distinction.

## 3. Figures and tables

*Strengths*
- Three figures (indexed raw quantities, product-level DAA wave, estimated courses) cover the main story.
- Tables give exact milestones, raw quantities, and course estimates.
- All figures and tables are cited in the body.

*Reviewer concerns / actions*
- **(Low)** Figure 1 overlays many time series (indexed) on one panel. Consider whether ribavirin, with a 20-fold rebound, compresses the other series. Current figure uses a right axis for total DAA, but indexed scales are hard to compare. The legend and caption explain the split, so this is acceptable.
- **(Low)** Figure 2 may be crowded; verify all product labels are legible at 300 dpi. Current build uses 2000×1200 px at 300 dpi; text size should be checked in the actual image.
- **(Fixed in this build)** Figures are supplied as separate high-resolution files; the submission manuscript contains legends only, per EID instructions. An inline review version is also provided.

## 4. Reproducibility and data provenance

*Strengths*
- All code and derived data are open at https://github.com/bougtoir/ndb-dea-drug-lag.
- NDB Open Data are publicly re-downloadable.
- Pipeline regenerates every number, figure, table, and reference access date from the data.

*Reviewer concerns / actions*
- **(High)** Ensure the public repo actually contains the EID branch changes. The `bougtoir/wip` sync workflow was updated to map `devin/1787736896-hcv-eid` to `ndb-dea-drug-lag`; merge of this PR will trigger the scheduled sync.
- **(Medium)** The `eid_word_counts.json` file and `eid_submission/` extracted directory are generated artifacts and should not be committed. Added to `.gitignore` in this build.
- **(Low)** Add the `make_submission_zip.py --journal eid` step to the README reproduce list. Done.

## 5. Strength of claims

*Strengths*
- Claims are consistently framed as descriptive and consistent with, rather than proof of, mechanisms.
- The "surge-then-decline" pattern is explicitly linked to curative therapy moving through a prevalent pool, with appropriate uncertainty language.

*Reviewer concerns / actions*
- **(Medium)** Abstract: "National hepatitis C treatment shifted almost entirely..." — because the metric is dispensing/estimated courses, consider "National hepatitis C antiviral dispensing shifted almost entirely..." or keep the current text if the Methods caveat is clear. The current abstract retains the "estimated treatment courses" qualifier, so the statement is supportable.
- **(Low)** "within one fiscal year of reimbursement" — this is based on the FY2014→FY2015 DAA share change. Because the first NHI listing occurred in FY2014 (September 2014), the one-year window is compressed. The text should make clear this is an annual-resolution observation, not a precise 12-month post-reimbursement estimate. The Discussion already notes annual resolution hides within-year timing.

## Priority summary

| Priority | Item | Status |
|---|---|---|
| Must fix before submission | Define DAA at first main-text use | Fixed |
| Must fix before submission | Separate figure files; no embedded figures in submission manuscript | Fixed |
| Must fix before submission | Sync mapping for public repo | Fixed |
| Must fix before submission | Ignore generated `eid_submission/` and `eid_word_counts.json` | Fixed |
| High | Keep pre-peak segment as non-interpretable; do not present as a trend | Already cautious; keep wording |
| High | Verify public repo reproduces outputs after merge | To be confirmed post-merge via sync-to-repos workflow |
| Medium | Soften cross-country comparison language | Already cautious; no change needed unless reviewer asks |
| Medium | Add brief justification for exponential decay choice | Optional; can add if reviewer requests |
| Low | Verify figure label legibility at 300 dpi | Check visually before upload |

## Final verdict

The manuscript meets EID's formal requirements (title page, 150-word unstructured abstract, ≤3,500 main text, ≤50 references, parenthetical italic numbered citations, separate figure files, tables after references, STROBE checklist, cover letter with word/reference counts). The substantive concerns above are manageable and framed transparently. The submission package is ready for upload once the public-repo sync is confirmed.
