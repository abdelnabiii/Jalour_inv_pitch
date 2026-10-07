# QA report: Green Square, ticket EGP 50m

## 1. Number tie-out (annex Summary vs deck, memo, DD memo)

| Figure | Annex cell | Annex value | Text searched | Found in | Result |
|---|---|---|---|---|---|
| Jalour net cash flow | Summary!B6 | 375.4475 | 375/375.4 | Deck: yes, Memo: yes | OK |
| NPV at 14% | Summary!B7 | 215.3114 | 215/215.3 | Deck: yes, Memo: yes | OK |
| Peak shortage | Summary!B8 | 102.5000 | 102/102.5/103 | Deck: yes, Memo: yes, DD: yes | OK |
| Total uses at peak | Summary!B14 | 112.9574 | 113/113.0 | Deck: yes, Memo: yes | OK |
| Balance of project funding | Summary!B15 | 62.9574 | 63/63.0 | Deck: yes, Memo: yes | OK |
| Ticket as % of modelled peak | Summary!B17 | 0.4878 | 48.8%/49% | Deck: yes, Memo: yes | OK |
| Option A face value | Summary!B20 | 100.0000 | 100/100.0 | Deck: yes, Memo: yes | OK |
| Option A value at final list | Summary!B22 | 143.3692 | 143/143.4 | Memo: yes | OK |
| Option A investor IRR base | Summary!B23 | 0.2640 | 26%/26.4% | Deck: yes, Memo: yes | OK |
| Option A investor MOIC base | Summary!B24 | 2.7814 | 2.78x/2.8x | Deck: yes, Memo: yes | OK |
| Option A IRR downside | Summary!B26 | 0.1864 | 18.6%/19% | Deck: yes, Memo: yes | OK |
| Option A MOIC downside | Summary!B27 | 2.5032 | 2.50x/2.5x | Deck: yes, Memo: yes | OK |
| Option A IRR upside | Summary!B28 | 0.2920 | 29%/29.2% | Deck: yes, Memo: yes | OK |
| Jalour cost of capital, Option A | Summary!B30 | 0.2074 | 20.7%/21% | Memo: yes, DD: yes | OK |
| Option A NPV cost to Jalour | Summary!B31 | -9.5465 | -10/-9.5 | Memo: yes, DD: yes | OK |
| Option B investor IRR | Summary!B34 | 0.2655 | 26.5%/27% | Deck: yes, Memo: yes, DD: yes | OK |
| Option B investor MOIC | Summary!B35 | 2.3000 | 2.30x/2.3x | Memo: yes | OK |
| Option B lowest cash-available coverage | Summary!B42 | 10.1033 | 10.10x/10.1x | Deck: yes, Memo: yes | OK |
| Option B quarters below 1.5x (net cash flow) | Summary!B40 | 6.0000 | 6 | Memo: yes, DD: yes | OK |
| Option B IRR if S2 | Summary!B45 | 0.2903 | 29%/29.0% | Memo: yes, DD: yes | OK |
| Option B IRR if S3 | Summary!B46 | 0.2784 | 27.8%/28% | Memo: yes, DD: yes | OK |
| Stress peak shortage | python | 269.1979 | 269/269.2 | Deck: yes, Memo: yes, DD: yes | OK |
| Downside peak shortage | python | 441.1946 | 441/441.2 | Memo: yes, DD: yes | OK |
| Ticket | python | 50.0000 | 50/50.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Option B total return | python | 115.0000 | 115/115.0 | Deck: yes, Memo: yes | OK |
| Instalment | python | 9.5833 | 10/9.58/9.6 | Memo: yes | OK |

Tie-out misses: 0

Annex inputs without a source or basis note: 0

Hard-coded numeric cells on calculation sheets (excluding blue scenario parameters): 0

## 2. Arithmetic check: independent recomputation (pure-Python bisection) vs spreadsheet

| Measure | Independent | Spreadsheet | Abs difference | Result |
|---|---|---|---|---|
| Option A investor IRR | 0.264030 | 0.264030 | 6.66e-16 | OK |
| Option A investor MOIC | 2.781362 | 2.781362 | 4.00e-15 | OK |
| Option A Jalour cost of capital | 0.207420 | 0.207420 | 2.28e-15 | OK |
| Option A Jalour NPV @14% | -9.546510 | -9.546510 | 2.49e-14 | OK |
| Option B S1 investor IRR | 0.265464 | 0.265464 | 5.00e-16 | OK |
| Project NPV @14% | 215.311393 | 215.311393 | 1.71e-13 | OK |
| Project net cash flow | 375.447528 | 375.447527 | 5.68e-14 | OK |
| Peak cumulative shortage | 102.500000 | 102.500000 | 0.00e+00 | OK |
| Option B min coverage (net cash flow) | -17.900956 | -17.900956 | 2.84e-14 | OK |
| Option B quarters below 1.5x | 6.000000 | 6.000000 | 0.00e+00 | OK |
| Option B min coverage (cash available) | 10.103323 | 10.103323 | 4.09e-14 | OK |

Python engine vs spreadsheet reconciliation sheet: ALL CHECKS ZERO

## 3. Disclosure scan (every file, including document properties, notes, hidden sheets and comments)

| File | Parts scanned | Prohibited hits | Bare "250" contexts |
|---|---|---|---|
| DD_Memo.docx | 18 | 0   | 0  |
| Deck.pptx | 98 | 0   | 0  |
| Financial_Annex.xlsx | 21 | 0   | 2 xl/worksheets/sheet10.xml: ...44406875    AE378+AF377  250.175367773437    AF378+AG... | xl/worksheets/sheet7.xml: ....741935483871    88      250    B7*1000/Inputs!$B$28... |
| Investment_Memo.docx | 18 | 0   | 2 word/document.xml: ...162.1                   250                   160... | word/document.xml: ...162.1                    250                    160... |

Prohibited-term hits: 0. Bare "250" occurrences are listed for context only: they are prices or the units face value, not a programme total.

## 4. Consistency of ticket, use of funds and returns across documents

- Ticket EGP 50: Deck yes, Memo yes, DD yes, annex yes (Summary links)
- Total uses: Deck yes, Memo yes, annex yes (Summary links)
- Option A IRR: Deck yes, Memo yes, annex yes (Summary links)
- Option B IRR: Deck yes, Memo yes, DD yes, annex yes (Summary links)

## 6. Challenge pass: five toughest questions from a sceptical family office

- **Why should list prices rise 33% to 48% when market growth is 3% to 16%?** Answered in: Memo 4.4 and 13.1; Deck slides 5 and 14; DD 5. Check: present
- **What if delivery payments arrive at handover, not before?** Answered in: Memo 11.1 and 13.2; Deck slide 11; DD red flag 1. Check: present
- **Who pays if there is a funding gap and why is there no guarantee?** Answered in: Memo 7, 11.1, 14.2; DD 10 and 11. Check: present
- **Does cash cover the Option B instalments?** Answered in: Memo 13.3; Deck slide 14; DD 8.2. Check: present
- **Can I get out of Option A early and what does Jalour need from the landlord?** Answered in: Memo 6.1, 12.1, 16.1; DD 2. Check: present

**Total failures requiring attention: 0**