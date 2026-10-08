# QA summary (INTERNAL)

Checks run on every pack: (1) number tie-out annex to deck, memo and DD memo; (2) independent recomputation of IRR, MOIC, NPV and coverage in pure Python versus the spreadsheet; (3) disclosure scan of every part of every file, including document properties, notes, hidden sheets and comments; (4) cross-document consistency; (5) visual render of decks and memos; (6) challenge pass with five sceptical-investor questions. Detailed reports: `QA_<pack>.md`.

| Pack | Slides | Annex formulas | Tie-out OK | Tie-out CHECK | Max abs difference, independent vs spreadsheet | Differences to explain | Prohibited-term hits | Annex reconciliation all zero | Failures |
|---|---|---|---|---|---|---|---|---|---|
| GS_050M | 17 | 20,873 | 37 | 0 | 1.7e-13 | 0 | 0 | yes | 0 |
| GS_075M | 17 | 20,873 | 37 | 0 | 1.7e-13 | 0 | 0 | yes | 0 |
| GS_100M | 17 | 20,873 | 37 | 0 | 1.7e-13 | 0 | 0 | yes | 0 |
| GS_125M | 18 | 22,486 | 42 | 0 | 1.7e-13 | 0 | 0 | yes | 0 |
| LA_050M | 17 | 20,873 | 37 | 0 | 6.8e-13 | 0 | 0 | yes | 0 |
| LA_075M | 17 | 20,873 | 37 | 0 | 6.8e-13 | 0 | 0 | yes | 0 |
| LA_100M | 17 | 20,873 | 37 | 0 | 6.8e-13 | 0 | 0 | yes | 0 |
| LA_125M | 18 | 22,486 | 42 | 0 | 6.8e-13 | 0 | 0 | yes | 0 |

Notes:
- The bare number "250" appears as the units face value (2.0 x 125) in the two 125M packs, as list prices (EGP thousand per m2) and in annex formulas; each context was reviewed. It never appears as a programme total.
- The two 125M packs deliberately show the name and finances of the other project (a Jalour decision); all other disclosure checks still apply to them, and the six other packs have zero mentions of the other project.
- Market data in the packs come from search summaries of broker reports and listings; they are cited in each memo (Appendix D) and need verification against the originals before issue.
- Visual inspection covered GS_100M (deck, memo, DD memo), LA_050M (deck), GS_125M (deck), LA_125M (memo); all decks and memos share the same templates.