# QA report: Green Square, ticket EGP 125m

## 1. Number tie-out (annex Summary vs deck, memo, DD memo)

| Figure | Annex cell | Annex value | Text searched | Found in | Result |
|---|---|---|---|---|---|
| Jalour net cash flow | Summary!B6 | 375.4475 | 375/375.4 | Deck: yes, Memo: yes | OK |
| NPV at 14% | Summary!B7 | 215.3114 | 215/215.3 | Deck: yes, Memo: yes | OK |
| Peak shortage | Summary!B8 | 102.5000 | 102/102.5/103 | Deck: yes, Memo: yes, DD: yes | OK |
| Total uses at peak | Summary!B14 | 112.9574 | 113/113.0 | Deck: yes, Memo: yes | OK |
| Balance of project funding | Summary!B15 | 0.0000 | 0/0.0 | Deck: yes, Memo: yes | OK |
| Ticket as % of modelled peak | Summary!B17 | 1.2195 | 122%/122.0% | Deck: yes, Memo: yes | OK |
| Option A face value | Summary!B20 | 250.0000 | 250/250.0 | Deck: yes, Memo: yes | OK |
| Option A value at final list | Summary!B22 | 358.4229 | 358/358.4 | Memo: yes | OK |
| Option A investor IRR base | Summary!B23 | 0.2640 | 26%/26.4% | Deck: yes, Memo: yes | OK |
| Option A investor MOIC base | Summary!B24 | 2.7814 | 2.78x/2.8x | Deck: yes, Memo: yes | OK |
| Option A IRR downside | Summary!B26 | 0.1864 | 18.6%/19% | Deck: yes, Memo: yes | OK |
| Option A MOIC downside | Summary!B27 | 2.5032 | 2.50x/2.5x | Deck: yes, Memo: yes | OK |
| Option A IRR upside | Summary!B28 | 0.2920 | 29%/29.2% | Deck: yes, Memo: yes | OK |
| Jalour cost of capital, Option A | Summary!B30 | 0.2074 | 20.7%/21% | Memo: yes, DD: yes | OK |
| Option A NPV cost to Jalour | Summary!B31 | -23.8663 | -23.9/-24 | Memo: yes, DD: yes | OK |
| Option B investor IRR | Summary!B34 | 0.2655 | 26.5%/27% | Deck: yes, Memo: yes, DD: yes | OK |
| Option B investor MOIC | Summary!B35 | 2.3000 | 2.30x/2.3x | Memo: yes | OK |
| Option B lowest cash-available coverage | Summary!B42 | 1.4102 | 1.41x/1.4x | Deck: yes, Memo: yes | OK |
| Option B quarters below 1.5x (net cash flow) | Summary!B40 | 9.0000 | 9 | Memo: yes, DD: yes | OK |
| Option B IRR if S2 | Summary!B45 | 0.2903 | 29%/29.0% | Memo: yes, DD: yes | OK |
| Option B IRR if S3 | Summary!B46 | 0.2784 | 27.8%/28% | Memo: yes, DD: yes | OK |
| Combined lowest cumulative position | Summary!B48 | 167.9546 | 168/168.0 | Deck: yes, Memo: yes | OK |
| Combined downside peak shortage | Summary!B51 | 1088.9735 | 1,089/1,089.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Combined peak, delivery payments at handover | Summary!B52 | 318.3333 | 318/318.3 | Deck: yes, Memo: yes | OK |
| Other project NPV | Summary!B54 | 307.5324 | 307.5/308 | Deck: yes, Memo: yes | OK |
| Other project net cash flow | Summary!B53 | 595.2929 | 595/595.3 | Deck: yes, Memo: yes | OK |
| Stress peak shortage | python | 269.1979 | 269/269.2 | Deck: yes, Memo: yes, DD: yes | OK |
| Downside peak shortage | python | 441.1946 | 441/441.2 | Memo: yes, DD: yes | OK |
| Ticket | python | 125.0000 | 125/125.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Option B total return | python | 287.5000 | 287.5/288 | Deck: yes, Memo: yes | OK |
| Instalment | python | 23.9583 | 23.96/24/24.0 | Memo: yes | OK |

Tie-out misses: 0

Annex inputs without a source or basis note: 0

Hard-coded numeric cells on calculation sheets (excluding blue scenario parameters): 0

## 2. Arithmetic check: independent recomputation (pure-Python bisection) vs spreadsheet

| Measure | Independent | Spreadsheet | Abs difference | Result |
|---|---|---|---|---|
| Option A investor IRR | 0.264030 | 0.264030 | 2.22e-16 | OK |
| Option A investor MOIC | 2.781362 | 2.781362 | 8.88e-16 | OK |
| Option A Jalour cost of capital | 0.207420 | 0.207420 | 2.28e-15 | OK |
| Option A Jalour NPV @14% | -23.866276 | -23.866276 | 5.68e-14 | OK |
| Option B S1 investor IRR | 0.265464 | 0.265464 | 9.44e-16 | OK |
| Project NPV @14% | 215.311393 | 215.311393 | 1.71e-13 | OK |
| Project net cash flow | 375.447528 | 375.447527 | 5.68e-14 | OK |
| Peak cumulative shortage | 102.500000 | 102.500000 | 0.00e+00 | OK |
| Option B min coverage (net cash flow) | -7.160383 | -7.160383 | 9.77e-15 | OK |
| Option B quarters below 1.5x | 9.000000 | 9.000000 | 0.00e+00 | OK |
| Option B min coverage (cash available) | 1.410216 | 1.410216 | 3.02e-14 | OK |

Python engine vs spreadsheet reconciliation sheet: ALL CHECKS ZERO

## 3. Disclosure scan (every file, including document properties, notes, hidden sheets and comments)

| File | Parts scanned | Prohibited hits | Bare "250" contexts |
|---|---|---|---|
| DD_Memo.docx | 18 | 0   | 0  |
| Deck.pptx | 104 | 0  Allowed in context: 250 million x1 in ppt/slides/slide2.xml (units face value 2.0 x 125); 250 million x1 in ppt/slides/slide3.xml (units face value 2.0 x 125); 250 million x1 in ppt/slides/slide13.xml (units face value 2.0 x 125) | 3 ppt/slides/slide2.xml: ...with a face value of EGP 250 million (2.0x the ticket... | ppt/slides/slide3.xml: ...with a face value of EGP 250 million at launch list p... | ppt/slides/slide13.xml: ...with a face value of EGP 250 million (2.0x) at launch... |
| Financial_Annex.xlsx | 22 | 0   | 5 xl/worksheets/sheet2.xml: ...32    Option_A!B5  250    B20/Inputs!$B$52  5.2... | xl/worksheets/sheet6.xml: ...1.21951219512195       250    MIN(B12,B8)/B8  1... | xl/worksheets/sheet10.xml: ...44406875    AE378+AF377  250.175367773437    AF378+AG... |
| Investment_Memo.docx | 18 | 0  Allowed in context: 250 million x3 in word/document.xml (units face value 2.0 x 125) | 6 word/document.xml: ...with a face value of EGP 250 million, twice the ticke... | word/document.xml: ...value                    250                       Op... | word/document.xml: ...162.1                   250                   160... |

Prohibited-term hits: 0. Bare "250" occurrences are listed for context only: they are prices or the units face value, not a programme total.

## 4. Consistency of ticket, use of funds and returns across documents

- Ticket EGP 125: Deck yes, Memo yes, DD yes, annex yes (Summary links)
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