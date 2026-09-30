# How fast did hepatitis C therapy switch? The interferon-to-DAA transition in Japan's NDB Open Data

Once a new therapy is approved and reimbursed, how quickly does a whole population
leave the standard therapy it replaces? A steep fall in national use of the prior
standard indicates that little friction stands between reimbursement and treatment;
a slow change would point to real practical friction (cost, referral pathways,
capacity, clinical caution).

We describe this in a setting where substitution was nearly complete: **chronic
hepatitis C**, where interferon (IFN)-based therapy was the standard of care until
interferon-free **direct-acting antivirals (DAAs)** were approved and listed for
National Health Insurance reimbursement in Japan in 2014-2015.

## Data
- **NDB Open Data** (National Database of Health Insurance Claims), editions 1-10,
  covering fiscal years **FY2014-FY2023**. Source: MHLW,
  https://www.mhlw.go.jp/ndb/opendatasite/
- Tables used: 処方薬 「性年齢別薬効分類別数量」 (内服/外用/注射 × 外来院内・外来院外・入院).
- Metric: **総計 (処方数量)** = national dispensed quantity per drug (tablets/capsules
  for oral drugs; pre-filled syringes/vials for injections). This is *not* a patient
  count; quantities are comparable within a product over time but are not additive
  across products with different dosage units.
- **Publication rules (censoring).** Each NDB table (formulation × care setting) lists
  only the highest-ranked products of each therapeutic class (top 30 in FY2014, top
  100 from FY2015; up to 500 in some tables), and any cell below **1,000** is shown as
  `-`. A product that is absent from a table or dashed is therefore *below the
  publication threshold*, not necessarily unused. `build_dataset.py` keeps a
  `suppressed` flag per row, records per-table listing caps and minimum listed totals
  (`data/ndb_publication_thresholds.csv`), and writes an **upper-bound** series
  (`data/hcv_timeseries_upper.csv`, `data/hcv_product_timeseries_upper.csv`,
  `data/censoring_bounds_long.csv`) in which every unpublished cell from a product's
  first listing onward is filled with the largest quantity it could conceal (the
  smallest published total of its class in that table when the listing cap was
  reached, otherwise 999). The primary series treats unpublished cells as zero and is
  a **lower bound**; the upper-bound series is a *bound*, not an observed total.

Drug groups (identified from the real drug names present in the files):
- **IFN-based therapy** — peginterferon (ペガシス, ペグイントロン) is the HCV-specific
  marker. Conventional interferon (スミフェロン, フエロン, イントロンA) is reported
  separately as a background series because it is not HCV-specific.
- **Ribavirin** (レベトール, コペガス) is reported separately and deliberately *not* as
  part of IFN-based therapy: it accompanied both peginterferon and the interferon-free
  sofosbuvir + ribavirin regimen for genotype 2, so its series mixes old and new
  treatment. It is excluded from course estimates.
- **Interferon-free DAAs** — 9 products: sofosbuvir, ledipasvir/sofosbuvir,
  daclatasvir, asunaprevir, glecaprevir/pibrentasvir, sofosbuvir/velpatasvir,
  elbasvir, grazoprevir, ombitasvir/paritaprevir/ritonavir.
- **First-generation NS3/4A protease inhibitors (IFN-based)** — simeprevir,
  telaprevir, vaniprevir. These were used *with* peginterferon+ribavirin (not
  interferon-free), so they are tabulated as a separate group (`PI_ifn`) and
  excluded from the interferon-free DAA total to keep that total clean.

## Headline findings
- Peginterferon dispensing fell **-99.2%** from FY2014 to FY2023, and the
  first-generation protease inhibitors given with it fell below the NDB publication
  threshold in every table from **FY2016**.
- Ribavirin fell too and was below the publication threshold from **FY2020**. Its
  FY2018-FY2019 published values come only from the tables in which it remained
  listed, and their movement lies within what the unlisted tables could conceal, so
  the apparent FY2019 uptick is a censoring artefact and is not interpreted.
- Total interferon-free DAA dispensing **peaked in FY2015** (+188% vs FY2014), then fell
  **-92%** to FY2023 (upper-bound series: about **-90%**; unpublished cells could
  conceal at most ~34% of the published DAA total in a single year, and change the
  DAA share of estimated courses by <0.2 percentage points; see
  `results/summary.json` → `censoring`, `results/its_summary.json` and
  `results/course_estimate.json` → `censoring_upper_bound`).
- On an estimated-course scale, interferon-free regimens went from **51%** to **97%** of
  hepatitis C antiviral treatment within one fiscal year of the first interferon-free
  listing; combined estimated volume peaked in FY2015 and then fell **~89%**.
- Interpretation: national dispensing switched within about a year of reimbursement.
  The surge-then-decay of total volume is the shape expected when a curative therapy
  works through a prevalent pool, and is consistent with — but not evidence for —
  accumulated demand. Media coverage and exposure to it were not measured.

Formal trend models with uncertainty intervals (`results/its_summary.json`): because
NDB begins in FY2014 (coincident with IFN-free DAA launch) there is no internal
pre-intervention baseline and a conventional pre/post interrupted time series is not
identifiable, so with n=10 annual points we fit descriptive trend models. Peginterferon
declined **~36%/yr** (95% CI ~24-47%), ribavirin **~76%/yr** (over FY2014-FY2019,
mixing both regimens and affected by censoring), conventional IFN **~25%/yr**;
total DAA dispensing fell **~25%/yr** after the FY2015 peak (segmented log-linear
regression, post- vs pre-peak slope change **P < 0.001**). The pre-peak segment rests on
only two observations and its bootstrap interval includes zero, so it is not interpreted
as a rate. Intervals are Newey-West (HAC) and residual-bootstrap based.

**Treatment-course estimates** (`results/course_estimate.json`,
`data/daa_course_assumptions.csv`, `data/ifn_course_assumptions.csv`): because dispensed
quantity is not a patient count and an 8-week tablet regimen is not comparable with a
48-week weekly injection, we convert quantities to approximate treatment courses (daily
dose x standard duration per the Japanese package inserts / JSH guideline), counting one
anchor product per two-drug regimen to avoid double-counting. This gives an estimated
peak of **~90,000 DAA courses in FY2015** and **~266,000-298,000 courses over
FY2014-FY2023** (duration-sensitivity range). Adding peginterferon courses (1 syringe =
1 weekly dose; 48-week course, 24 weeks as sensitivity) gives a combined national
treatment volume and the DAA share of it. Protease inhibitors are not counted separately
because they were always added to peginterferon. These are explicit estimates, not
observed patient counts.

All numbers above are regenerated into `results/summary.json`,
`results/its_summary.json` and `results/course_estimate.json`; do not hard-code.
Literature-derived context numbers used in the Introduction/Discussion (Japan carrier
estimates, Polaris cascade-of-care, Australia/Taiwan program figures, US treatment
uptake) live in `data/discussion_support.json` with the reference key for each value.

## "News / announcement" side (`data/announcement_events.csv`)
Milestones are official **NHI drug-price listings (薬価収載)** and **PMDA approvals**:
1. **daclatasvir + asunaprevir** — approval **2014-07-04**, NHI listing **2014-09-02**
   (launch 2014-09-03); world-first all-oral, IFN/ribavirin-free regimen for chronic
   hepatitis C (Bristol-Myers Squibb).
2. **sofosbuvir (Sovaldi)** — approval **2015-03-26**, NHI listing **2015-05-20**;
   **ledipasvir/sofosbuvir (Harvoni)** — approval **2015-07-03**, NHI listing
   **2015-08-31**.
3. **glecaprevir/pibrentasvir (Maviret)** — approval **2017-09-27**, NHI listing
   **2017-11-22** (launch 2017-11-27), pangenotypic.

Dates are **day-precise**, verified against the products' Japanese package inserts /
interview forms (承認年月日, 薬価基準収載日) and Chuikyo listing records, plus company
press releases; per-event sources are recorded in `data/announcement_events.csv`.
Unsourced clinical claims (e.g. cure rates) are deliberately excluded until a citable
source is documented.

## Reproduce
```bash
python3 scripts/download_ndb.py      # fetch raw NDB workbooks into data/ndb_raw/
python3 scripts/build_dataset.py     # -> data/target_drugs_long.csv (+suppressed flag), hcv_timeseries.csv,
                                     #    hcv_product_timeseries.csv, *_upper.csv (censoring upper bound),
                                     #    ndb_publication_thresholds.csv, censoring_bounds_long.csv
python3 scripts/analyze.py           # -> results/summary.json
python3 scripts/course_estimate.py   # -> results/course_estimate.json (treatment-course sensitivity, ESTIMATE)
python3 scripts/its_analysis.py      # -> results/its_summary.json (trend models + 95% intervals)
python3 scripts/make_figures.py --lang en
python3 scripts/make_figures.py --lang ja
python3 scripts/make_manuscript.py   # -> output/ JA+EN manuscript docx, P&DS, Hepatology Research,
                                     #    Emerging Infectious Diseases, Journal of Viral Hepatitis, and
                                     #    Journal of Gastroenterology and Hepatology variants, table docx,
                                     #    figure pptx, STROBE checklist, and cover letters
python3 scripts/make_manuscript.py --only jgh   # regenerate one journal variant only (other docx untouched)
python3 scripts/make_submission_zip.py --journal pds     # -> output/pds_submission.zip
python3 scripts/make_submission_zip.py --journal hepres  # -> output/hepres_submission.zip
python3 scripts/make_submission_zip.py --journal eid     # -> output/eid_submission.zip
python3 scripts/make_submission_zip.py --journal jvh     # -> output/jvh_submission.zip
python3 scripts/make_submission_zip.py --journal jgh     # -> output/jgh_submission.zip (JGH, Wiley transfer)
```

Model equations in the manuscript are written as native Word equations (Office Math /
OMML), not LaTeX, so they render and remain editable in Microsoft Word.

## Data and code availability
All analysis code, derived datasets, and the figure/manuscript generators are openly
available at **https://github.com/bougtoir/ndb-dea-drug-lag**. The raw NDB Open Data
workbooks are published by MHLW and are re-downloadable with `scripts/download_ndb.py`,
so the full pipeline reproduces every reported number, figure, and table from the
public data.

## Limitations
- NDB Open Data begins in FY2014, the same period IFN-free DAAs launched, so there is
  no pre-DAA IFN baseline *within* NDB; the FY2014 value already reflects decline from
  the pre-2014 peak.
- The metric is dispensed quantity, not patient counts; units differ across products,
  so summed DAA quantity is not a patient count.
- NDB Open Data lists only top-ranked products and masks cells below 1,000, so
  low-volume products and the tail of a declining drug are underestimated; apparent
  disappearance means only "below the publication threshold". The upper-bound
  sensitivity series quantifies how much this could matter (see Data).
- Annual resolution only (no within-year interrupted time series); reported trend
  rates carry wide uncertainty intervals given the small number of annual
  observations (n=10), and there is no control condition or placebo event.
- Descriptive evidence only: no measure of media coverage or of individual treatment
  decisions was available, and nothing here identifies why any individual was treated
  when they were.
- Ribavirin cannot be assigned to either regimen, and conventional interferon is not
  HCV-specific; both are reported as separate series.
- Secular changes from FY2020 onward, including the COVID-19 pandemic's effect on
  outpatient visits and prescribing, may also have influenced later dispensing and
  cannot be separated from the ongoing DAA decline.
- Treatment-course figures are estimates dependent on regimen/duration assumptions
  (`data/daa_course_assumptions.csv`), not observed patient counts.
