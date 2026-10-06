# PDS Requirements Audit — Pharmacoepidemiology and Drug Safety (Wiley)

Source checked: Wiley author page for *Pharmacoepidemiology and Drug Safety*
(onlinelibrary.wiley.com, "Author Guidelines" / "For authors") and the Wiley
"Free Format submission" policy, checked 2026-10-06 UTC.
Where the handoff memory conflicted with current official instructions, the
official instructions were followed and are noted below.

## Verified current requirements and compliance

| Requirement (current official) | This package |
|---|---|
| Free-format submission: any consistent, readable style accepted at first submission | Times New Roman 12 pt, 1.5 spacing, 25 mm margins, page numbers — readable Wiley-style default |
| Original Article: ~3,000 words main text (abstract, tables, figure captions, references excluded) | Body ≈ 2,986 words excluding the declaration statements (3,121 including them) — at/under the guideline |
| Structured abstract ≤ 250 words: Purpose / Methods / Results / Conclusions | Purpose–Methods–Results–Conclusions, 234 words |
| ≤ 7 keywords | 7 keywords |
| Up to 5 "Key Points" bullet statements | 5 Key Points |
| Plain Language Summary ≤ 200 words, single paragraph, jargon-free | 190 words, one paragraph |
| References: consistent style acceptable (free format) | Vancouver, numbered in order of first appearance, superscript in-text |
| Figures/tables may be embedded near text or at end | Embedded after first citation; tables also supplied as separate editable docx; figures supplied as PNG + editable PPTX |
| ORCID iD required for corresponding author at submission | Placeholder `[ORCID]` on title page — author completes |
| Title page: article type, running head, corresponding author, ethics/funding/COI statements | Present (`add_pds_titlepage`, `add_pds_statements`) |
| PDS-specific conflict-of-interest form | Flagged in FINAL_HANDOFF_PDS.md — author downloads/signs at submission (cannot be generated here) |
| Transparency/reproducibility statement encouraged (Wang & Pottegård tradition) | Present: code/data URL + "no preregistered protocol" wording |
| Submission via Wiley Research Exchange portal | Human step, outside this package |

## Resolution of memory-vs-official conflicts

- Handoff memory said "original research reports ≤4,000 words". Current PDS
  guidance for Original Articles is ~3,000 words of main text. The official
  number won: the Discussion and Limitations were compressed so the body is
  ≈3,000 words.
- Structured abstract headings differ across Wiley journals; the PDS
  convention *Purpose / Methods / Results / Conclusions* was used (not
  Background/Objectives).
- Citation style is free-format at submission; Vancouver superscript kept for
  consistency with the earlier journal variants and PDS house convention.

## Items that cannot be verified mechanically

- Exact wording of the PDS COI form (author downloads it in the portal).
- Graphical abstract is optional — not produced (not required).
- Reviewer suggestions (≤3) left as a placeholder in the cover letter.
