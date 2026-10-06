# PDS Hostile Review — pre-submission self-audit

Each question is the strongest version of the expected objection, followed by
where the manuscript answers it and any residual risk.

1. **"This is an ecological description of dispensing counts, not pharmacoepidemiology."**
   Answer: the manuscript never claims patient-level inference. It frames the
   unit explicitly (dispensed quantity → estimated courses, "not observed
   patient counts"), uses the standard drug-utilization common-scale logic
   (ref 1), and is honest that all trend models are descriptive (Limitations).
   Residual risk: low-moderate. Some reviewers may still want claims data; the
   scope statement in the Introduction ("aggregate open data") preempts the
   bait-and-switch accusation but cannot add individual-level evidence.

2. **"The course-equivalent conversion is assumption-stacked — a 48-week peginterferon course, an 8–12-week DAA course, one anchor product per regimen."**
   Answer: the conversion is stated fully and transparently in Methods, and a
   shorter-duration sensitivity analysis quantifies the dependence (course
   totals 270,939–305,499). The Discussion acknowledges estimates "track those
   counts in scale and peak timing" vs claims data — an external cross-check,
   not blind faith.

3. **"NDB Open Data masks small cells and drops low-ranked products — your 99.3% decline and 'below threshold' claims are artifacts of the publication threshold."**
   Answer: addressed at length in Limitations and by the dedicated upper-bound
   analysis: even under worst-case masking the post-peak decline moves only
   23.4%→21.7%/yr, the share <0.5pp, peak year unchanged. "Below the
   publication threshold" is stated, never "zero use". First-gen PIs also fell
   99%+ in published units before masking could matter (826,748→8,961 by FY2016
   alone).

4. **"You fit a segmented model with a knot placed at the observed peak and only 2 pre-peak points — that is not a changepoint."**
   Answer: stated in Methods ("model-based and exploratory"), in Results
   ("model-based"), and in Limitations ("slope-change P values are
   model-based rather than formal change-point evidence"). Reported "for
   completeness rather than as an estimated rate". P values are retained (per
   mandate) but never presented as independent changepoint evidence.

5. **"Annual data, n=11, no control series — your time-series analysis is overinterpretation."**
   Answer: the paper calls the models "descriptive" throughout, states n=11
   and "no control condition" in Limitations, and uses Newey–West +
   block-bootstrap CIs rather than pretending at causal ITS claims.

6. **"The post-peak decline could be the COVID-19 pandemic, not a depleted prevalent pool."**
   Answer: Limitations name FY2020 secular changes explicitly. The claim is
   calibrated: "a pattern consistent with progressive treatment of a prevalent
   population" — consistency, not proof — and the magnitude is anchored
   externally to ~520,000 in-care + 167,000–767,000 untreated carriers
   (Tanaka 2018).

7. **"Yamashita (claims, 357,877 patients) and Suoh (~300,000, ¥1.23tn) already did this."**
   Answer: differentiated in Introduction AND Discussion: those studies
   quantify *who was treated, where, and at what cost* on the successor series
   alone; this study includes the interferon-era baseline and measures
   replacement dynamics (incumbent exit speed, post-substitution throughput)
   on a common scale. Course totals are cross-checked against their patient
   counts and track them.

8. **"The decline of peginterferon might just be interferon's global obsolescence, not substitution."**
   Answer: Results include the background comparison — conventional interferon
   (non-HCV-specific) fell only 94.9% overall / 25% per year vs peginterferon's
   34% per year, so the HCV-specific products fell much faster than the
   general interferon trend. The Methods milestone table ties the collapse to
   specific approval/listing dates.

9. **"Ribavirin and dual-class products make your categories leaky."**
   Answer: Limitations: ribavirin "cannot be assigned to either therapy and is
   excluded from course estimates"; conventional interferon flagged as not
   HCV-specific. Anchor-product convention documented in Methods.

10. **"The 'rapid → slow' two-phase claim is curve-fitting a narrative onto one series."**
    Answer: the two phases correspond to observable, independently dated
    events (NHI listing dates; measured shares; measured combined-volume
    peak), and the interpretation is offered as consistency with a mechanism,
    externally anchored, not as fitted evidence of it.

11. **"Numbers inside a narrative claim — e.g. '97.0% within one fiscal year' — depend on the course conversion the reviewer may reject."**
    Answer: every converted quantity is labeled "estimated"; raw dispensed-
    quantity results (peginterferon 99.3%, conventional IFN 25%/yr, PI collapse)
    stand on the published units alone, so the headline substitution finding
    survives even if a reviewer discards the course scale.

12. **"Aggregate, non-stratified data — you cannot address confounding, genotype mix, or presciber behavior."**
    Answer: Limitations state the data "cannot be stratified by genotype,
    fibrosis stage, treatment history, or region" and that mechanisms
    (tolerability, listing, subsidy) were not measured individually. The
    Discussion attributes the speed to the combination explicitly rather than
    to any single unmeasured factor.

## Residual unfixable risks (declared, not papered over)

- No pre-DAA interferon baseline exists inside NDB Open Data (it starts FY2014);
  the analysis "speaks to the speed of the observed displacement rather than
  its full extent" — stated verbatim.
- Course-equivalent totals will read as somewhat low vs claims counts
  (anchored, disclosed, cross-checked — but a reviewer preferring the claims
  denominator will still prefer it).
