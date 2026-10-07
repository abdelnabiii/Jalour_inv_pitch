# QA report: L'avenir, ticket EGP 50m

## 1. Number tie-out (annex Summary vs deck, memo, DD memo)

| Figure | Annex cell | Annex value | Text searched | Found in | Result |
|---|---|---|---|---|---|
| Jalour net cash flow | Summary!B6 | 595.2929 | 595/595.3 | Deck: yes, Memo: yes | OK |
| NPV at 14% | Summary!B7 | 307.5324 | 307.5/308 | Deck: yes, Memo: yes | OK |
| Peak shortage | Summary!B8 | 102.0181 | 102/102.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Total uses at peak | Summary!B14 | 112.4755 | 112/112.5 | Deck: yes, Memo: yes | OK |
| Balance of project funding | Summary!B15 | 62.4755 | 62/62.5 | Deck: yes, Memo: yes | OK |
| Ticket as % of modelled peak | Summary!B17 | 0.4901 | 49%/49.0% | Deck: yes, Memo: yes | OK |
| Option A face value | Summary!B20 | 100.0000 | 100/100.0 | Deck: yes, Memo: yes | OK |
| Option A value at final list | Summary!B22 | 148.1481 | 148/148.1 | Memo: yes | OK |
| Option A investor IRR base | Summary!B23 | 0.2925 | 29%/29.2% | Deck: yes, Memo: yes | OK |
| Option A investor MOIC base | Summary!B24 | 2.8741 | 2.87x/2.9x | Deck: yes, Memo: yes | OK |
| Option A IRR downside | Summary!B26 | 0.2041 | 20%/20.4% | Deck: yes, Memo: yes | OK |
| Option A MOIC downside | Summary!B27 | 2.5867 | 2.59x/2.6x | Deck: yes, Memo: yes | OK |
| Option A IRR upside | Summary!B28 | 0.3228 | 32%/32.3% | Deck: yes, Memo: yes | OK |
| Jalour cost of capital, Option A | Summary!B30 | 0.2256 | 22.6%/23% | Memo: yes, DD: yes | OK |
| Option A NPV cost to Jalour | Summary!B31 | -10.4013 | -10/-10.4 | Memo: yes, DD: yes | OK |
| Option B investor IRR | Summary!B34 | 0.2655 | 26.5%/27% | Deck: yes, Memo: yes, DD: yes | OK |
| Option B investor MOIC | Summary!B35 | 2.3000 | 2.30x/2.3x | Memo: yes | OK |
| Option B lowest cash-available coverage | Summary!B42 | 1.7846 | 1.78x/1.8x | Deck: yes, Memo: yes | OK |
| Option B quarters below 1.5x (net cash flow) | Summary!B40 | 3.0000 | 3 | Memo: yes, DD: yes | OK |
| Option B IRR if S2 | Summary!B45 | 0.2866 | 28.7%/29% | Memo: yes, DD: yes | OK |
| Option B IRR if S3 | Summary!B46 | 0.2738 | 27%/27.4% | Memo: yes, DD: yes | OK |
| Stress peak shortage | python | 244.3073 | 244/244.3 | Deck: yes, Memo: yes, DD: yes | OK |
| Downside peak shortage | python | 848.6043 | 848.6/849 | Memo: yes, DD: yes | OK |
| Ticket | python | 50.0000 | 50/50.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Option B total return | python | 115.0000 | 115/115.0 | Deck: yes, Memo: yes | OK |
| Instalment | python | 9.5833 | 10/9.58/9.6 | Memo: yes | OK |

Tie-out misses: 0

Annex inputs without a source or basis note: 0

Hard-coded numeric cells on calculation sheets (excluding blue scenario parameters): 0

## 2. Arithmetic check: independent recomputation (pure-Python bisection) vs spreadsheet

| Measure | Independent | Spreadsheet | Abs difference | Result |
|---|---|---|---|---|
| Option A investor IRR | 0.292473 | 0.292473 | 1.62e-13 | OK |
| Option A investor MOIC | 2.874074 | 2.874074 | 2.22e-15 | OK |
| Option A Jalour cost of capital | 0.225574 | 0.225574 | 4.26e-14 | OK |
| Option A Jalour NPV @14% | -10.401255 | -10.401255 | 3.73e-14 | OK |
| Option B S1 investor IRR | 0.265464 | 0.265464 | 5.00e-16 | OK |
| Project NPV @14% | 307.532371 | 307.532371 | 6.82e-13 | OK |
| Project net cash flow | 595.292929 | 595.292929 | 1.14e-13 | OK |
| Peak cumulative shortage | 102.018063 | 102.018062 | 1.42e-14 | OK |
| Option B min coverage (net cash flow) | -13.485259 | -13.485259 | 3.38e-14 | OK |
| Option B quarters below 1.5x | 3.000000 | 3.000000 | 0.00e+00 | OK |
| Option B min coverage (cash available) | 1.784595 | 1.784595 | 1.78e-15 | OK |

Python engine vs spreadsheet reconciliation sheet: ALL CHECKS ZERO

## 3. Disclosure scan (every file, including document properties, notes, hidden sheets and comments)

| File | Parts scanned | Prohibited hits | Bare "250" contexts |
|---|---|---|---|
| DD_Memo.docx | 18 | 0   | 2 word/document.xml: ...00                   191,250                   106,00... | word/document.xml: ...etail                    250,000                    3... |
| Deck.pptx | 98 | 0   | 3 ppt/slides/slide5.xml: ...Launch EGP 250,000 per m2, rising to 36... | ppt/slides/slide9.xml: ...250... | ppt/charts/chart3.xml: ...General          250.488    250.488    187.86... |
| Financial_Annex.xlsx | 21 | 0   | 26 xl/worksheets/sheet3.xml: ...127      128      129    250    130    131      132... | xl/worksheets/sheet10.xml: ...nputs!$B$7,0)+AO17-AO20  250.8875435    AO22+IF(Sourc... | xl/worksheets/sheet10.xml: ...nputs!$B$7,0)+AP17-AP20  250.8875435    AP22+IF(Sourc... |
| Investment_Memo.docx | 18 | 0   | 9 word/document.xml: ...launch list (retail EGP 250,000 per m2, offices EGP... | word/document.xml: ...t prices rising from EGP 250,000 to 360,000 per m2 fo... | word/document.xml: ...at a launch list of EGP 250,000 (retail) and 160,000... |

Prohibited-term hits: 0. Bare "250" occurrences are listed for context only: they are prices or the units face value, not a programme total.

## 4. Consistency of ticket, use of funds and returns across documents

- Ticket EGP 50: Deck yes, Memo yes, DD yes, annex yes (Summary links)
- Total uses: Deck yes, Memo yes, annex yes (Summary links)
- Option A IRR: Deck yes, Memo yes, annex yes (Summary links)
- Option B IRR: Deck yes, Memo yes, DD yes, annex yes (Summary links)

## 6. Challenge pass: five toughest questions from a sceptical family office

- **Why should list prices rise 44% to 50% when market growth is 3% to 16%?** Answered in: Memo 4.4 and 13.1; Deck slides 5 and 14; DD 5. Check: present
- **What if delivery payments arrive at handover, not before?** Answered in: Memo 11.1 and 13.2; Deck slide 11; DD red flag 1. Check: present
- **Who pays if there is a funding gap and why is there no guarantee?** Answered in: Memo 7, 11.1, 14.2; DD 10 and 11. Check: present
- **Does cash cover the Option B instalments?** Answered in: Memo 13.3; Deck slide 14; DD 8.2. Check: present
- **Can I get out of Option A early and what does Jalour need from the landlord?** Answered in: Memo 6.1, 12.1, 16.1; DD 2. Check: present

**Total failures requiring attention: 0**