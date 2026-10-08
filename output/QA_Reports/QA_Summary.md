# QA summary (INTERNAL)

Checks run on every pack: (1) number tie-out annex to deck, memo and DD memo; (2) independent recomputation of IRR, MOIC, NPV and coverage in pure Python versus the spreadsheet; (3) disclosure scan of every part of every file, including document properties, notes, hidden sheets and comments; (4) cross-document consistency; (5) visual render of decks and memos; (6) challenge pass with five sceptical-investor questions. Detailed reports: `QA_<pack>.md`.

| Pack | Slides | Annex formulas | Tie-out OK | Tie-out CHECK | Max abs difference, independent vs spreadsheet | Differences to explain | Prohibited-term hits | Annex reconciliation all zero | Failures |
|---|---|---|---|---|---|---|---|---|---|
| GS_050M | 17 | 20,873 | 37 | 0 | 1.7e-13 | 0 | 0 | yes | 0 |
| GS_075M | 17 | 20,873 | 37 | 0 | 1.7e-13 | 0 | 0 | yes | 0 |
| GS_100M | 17 | 20,873 | 37 | 0 | 1.7e-13 | 0 | 0 | yes | 0 |
| LA_050M | 17 | 20,873 | 37 | 0 | 6.8e-13 | 0 | 0 | yes | 0 |
| LA_075M | 17 | 20,873 | 37 | 0 | 6.8e-13 | 0 | 0 | yes | 0 |
| LA_100M | 17 | 20,873 | 37 | 0 | 6.8e-13 | 0 | 0 | yes | 0 |

Notes:
- Six single-project packs (50, 75, 100M): each names only its own project; the disclosure scan forbids the other project name and every other prohibited term. The bare number "250" appears only as list prices (EGP thousand per m2) and in annex formulas, never as a programme total.
- COMBINED_125M (EGP 62.5M in each project, the only 125M case) is checked separately in QA_COMBINED_125M.md: tie-out, independent recomputation, disclosure scan (both project names allowed there), confidentiality statements, NDA template content.
- Every pack contains NDA_Template.docx; confidentiality statements appear on the cover and notice of every document and in the annex cover.

Combined pack COMBINED_125M: failures 0; annex reconciliation ALL CHECKS ZERO