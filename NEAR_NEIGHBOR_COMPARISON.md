# NEAR_NEIGHBOR_COMPARISON — Suoh et al. 2026 vs this manuscript

## Bibliographic verification
Suoh M, Esmaili S, Eslam M, George J. *Nationwide Socioeconomic Evaluation of Direct-Acting Antivirals for Chronic Hepatitis C Virus Infection in Japan Between 2014 and 2023.* J Viral Hepat. 2026;33(8):e70206. doi:10.1111/jvh.70206 (online 2026-07-15; issue Aug 2026). Confirmed via Wiley/Ovid TOC + WSLHD repository + abstract text.

## Suoh et al. — what they did (from verified abstract)
- Data: publicly available records incl. hepatitis **subsidy program** + universal healthcare claims (NDB-related public records), FY2014–2023.
- Unit: estimated treated people per DAA regimen (standard dose conversion) + drug cost.
- Endpoints: cumulative treated (~300,000), cumulative cost ≥1.23 trillion JPY; regimen-level prescription timing vs approval; indication; age shifts; **prefecture-level** prevalence/treated counts.
- Did NOT jointly model interferon-based therapy (abstract: DAAs only).
- Did NOT quantify therapeutic substitution (no incumbent-vs-successor exchange).
- Course-equivalent scale: partial — they convert DAA quantities to treated people per regimen, but never place peginterferon/IFN on the same scale.
- Transition dynamics: noted peak 2015 + declining trend; no rate/shape quantification reported.
- Post-peak decline interpreted as uptake pattern; elimination implication mentioned briefly ("suggests an active program of case finding and treatment").

## Current manuscript — what we do
- Data: NDB Open Data editions 1–10, national dispensed **quantities** FY2014–2023.
- Both therapy eras on one **course-equivalent scale**: peginterferon, first-gen PIs, conventional IFN, ribavirin AND all DAA products → estimated courses.
- Quantifies **substitution dynamics**: speed of IFN collapse (−99.2%, −36%/yr), IFN-free share 50.9%→97.0% within 1 FY, post-peak volume trajectory (−24.7%/yr descriptive segmented model).
- Post-peak decline interpreted explicitly as transition from prevalent-pool treatment to a lower-flow phase.
- No prefecture breakdown, no cost, no patient demographics — different axes from Suoh.

## Overlap
- Same country, same period, same DAA products, same order-of-magnitude treated volume (~300k).
- Both note the FY2015 peak and subsequent decline.

## Non-overlap
- IFN legacy therapy modeled jointly (Suoh: absent).
- Substitution speed/share quantified (Suoh: absent).
- Common course scale across eras (Suoh: DAAs only).
- Suoh adds: costs, prefectures, age/indication mix, subsidy-program records — we do not.

## Defensible novelty remaining
"Extends prior national utilization work by quantifying the **replacement dynamics** between incumbent interferon-based therapy and interferon-free DAAs on a common course-equivalent scale, and by characterizing the post-peak transition to a lower-flow treatment phase."

## Wording to avoid
- "first nationwide study of DAA use/uptake in Japan" — FALSE after Suoh 2026.
- "first national analysis of treatment volume" — FALSE (Suoh estimated treated counts).
- Implying no prior national DAA utilization study exists.
- Safe: "first to quantify the substitution itself on a course-equivalent scale" — Suoh does not do this; phrased as "extends" rather than "first" per instruction.

## Strongest defensible novelty statement
"Whereas Suoh et al. estimated the national number treated with DAAs, their cost, and their geographic/demographic pattern, the present study quantifies the **dynamic replacement** of the interferon-based standard by interferon-free regimens — the speed and completeness of the switch and the volume trajectory after substitution — a quantity no prior national study reports."
