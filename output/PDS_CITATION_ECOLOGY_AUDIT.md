# PDS Citation-Ecology Audit

Purpose: confirm the reference list places this study inside the
pharmacoepidemiology literature PDS reviewers expect — drug-utilization
methodology, national dispensing analyses, and the hepatitis C treatment
transition — without citation-gaming (no gratuitous PDS self-citations, no
padding).

## Reference-neighborhood check (18 references)

| Neighborhood | References in this manuscript | Assessment |
|---|---|---|
| Drug-utilization common scale | WHO ATC/DDD guidelines (ref 1) | Anchors the course-equivalent conversion in the field's standard common-scale logic |
| ITS / segmented-regression methods for medicine use | Wagner et al. 2002 *J Clin Pharm Ther*; Bernal et al. 2017 *Int J Epidemiol* (refs 11–12) | The two canonical methods citations a PDS reviewer would ask for when segmented trend models appear |
| Statistical machinery | Newey & West 1987; Efron & Tibshirani 1993 (refs 13–14) | Correct HAC + bootstrap citations, not decorative |
| Japanese HCV clinical/regulatory context | JSH guideline, BMS approval press release, Kumada, Omata, Mizokami phase-3 trials (refs 2–6) | Same clinical backbone the prior journal variants used; appropriate and verified |
| Prior national-scale work on the same transition | Yamashita et al. 2025 *Int J Infect Dis*; Suoh et al. 2026 *J Viral Hepat* (refs 7–8) | Cited in both Introduction and Discussion with the explicit differentiation: they quantify *who was treated, where, at what cost*; this study measures *replacement dynamics* on a course-equivalent scale. Framing is complementary, not adversarial |
| Data source + reimbursement milestones | MHLW NDB Open Data; Chuikyo NHI listing records (refs 9–10) | Primary sources, not secondary summaries |
| Prevalent-pool anchor | Tanaka et al. 2018 (ref 15) | Independent estimate of the HCV-in-care population that bounds the course-equivalent interpretation |
| Access/subsidy context | MHLW subsidy program; Setoyama et al. 2021 (refs 16–17) | Explains why Japanese substitution speed is not transferable without qualification |
| International contrast | Thompson et al. 2022 MMWR (ref 18) | US insured-population treatment fraction — the contrast a pharmacoepidemiology reviewer expects |

## Findings

- Every reference carries weight (method, data, or differentiation); none were
  added to signal journal fit.
- No PDS house-organ self-citations were added. The two methods papers are the
  natural citations for segmented regression in medication-use research and
  would appear on a PDS referee's checklist regardless of venue.
- Suoh et al. 2026 and Yamashita et al. 2025 are the only prior
  nationwide-scale studies of this transition and are differentiated
  explicitly, per the handoff requirement.
- Citation order is strictly first-appearance; verified against the built
  docx (refs 1–18 in sequence).
- No fabricated citations: all references existed in the validated `_REF_COMMON`
  set or were verified (Wagner DOI 10.1046/j.1365-2710.2002.00430.x; Bernal DOI
  10.1093/ije/dyw098; WHO ATC/DDD 2025, 28th ed.).

## Residual judgment calls

- A classical drug-utilization paper might also cite a *defined daily dose*
  methodology debate paper; omitted deliberately — the manuscript's claim is
  course-equivalence logic ("follows the same logic as"), which the WHO
  guideline itself suffices for.
- No citation claims generalizability beyond a methodological pattern; the
  transferability language stays in prose, not in cited authority.
