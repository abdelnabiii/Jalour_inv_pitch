# QA report: L'avenir, ticket EGP 125m

## 1. Number tie-out (annex Summary vs deck, memo, DD memo)

| Figure | Annex cell | Annex value | Text searched | Found in | Result |
|---|---|---|---|---|---|
| Jalour net cash flow | Summary!B6 | 595.2929 | 595/595.3 | Deck: yes, Memo: yes | OK |
| NPV at 14% | Summary!B7 | 307.5324 | 307.5/308 | Deck: yes, Memo: yes | OK |
| Peak shortage | Summary!B8 | 102.0181 | 102/102.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Total uses at peak | Summary!B14 | 112.4755 | 112/112.5 | Deck: yes, Memo: yes | OK |
| Balance of project funding | Summary!B15 | 0.0000 | 0/0.0 | Deck: yes, Memo: yes | OK |
| Ticket as % of modelled peak | Summary!B17 | 1.2253 | 122.5%/123% | Deck: yes, Memo: yes | OK |
| Option A face value | Summary!B20 | 250.0000 | 250/250.0 | Deck: yes, Memo: yes | OK |
| Option A value at final list | Summary!B22 | 370.3704 | 370/370.4 | Memo: yes | OK |
| Option A investor IRR base | Summary!B23 | 0.2925 | 29%/29.2% | Deck: yes, Memo: yes | OK |
| Option A investor MOIC base | Summary!B24 | 2.8741 | 2.87x/2.9x | Deck: yes, Memo: yes | OK |
| Option A IRR downside | Summary!B26 | 0.2041 | 20%/20.4% | Deck: yes, Memo: yes | OK |
| Option A MOIC downside | Summary!B27 | 2.5867 | 2.59x/2.6x | Deck: yes, Memo: yes | OK |
| Option A IRR upside | Summary!B28 | 0.3228 | 32%/32.3% | Deck: yes, Memo: yes | OK |
| Jalour cost of capital, Option A | Summary!B30 | 0.2256 | 22.6%/23% | Memo: yes, DD: yes | OK |
| Option A NPV cost to Jalour | Summary!B31 | -26.0031 | -26/-26.0 | Memo: yes, DD: yes | OK |
| Option B investor IRR | Summary!B34 | 0.2655 | 26.5%/27% | Deck: yes, Memo: yes, DD: yes | OK |
| Option B investor MOIC | Summary!B35 | 2.3000 | 2.30x/2.3x | Memo: yes | OK |
| Option B lowest cash-available coverage | Summary!B42 | 3.8443 | 3.84x/3.8x | Deck: yes, Memo: yes | OK |
| Option B quarters below 1.5x (net cash flow) | Summary!B40 | 5.0000 | 5 | Memo: yes, DD: yes | OK |
| Option B IRR if S2 | Summary!B45 | 0.2866 | 28.7%/29% | Memo: yes, DD: yes | OK |
| Option B IRR if S3 | Summary!B46 | 0.2738 | 27%/27.4% | Memo: yes, DD: yes | OK |
| Combined lowest cumulative position | Summary!B48 | 167.9546 | 168/168.0 | Deck: yes, Memo: yes | OK |
| Combined downside peak shortage | Summary!B51 | 1088.9735 | 1,089/1,089.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Combined peak, delivery payments at handover | Summary!B52 | 318.3333 | 318/318.3 | Deck: yes, Memo: yes | OK |
| Other project NPV | Summary!B54 | 215.3114 | 215/215.3 | Deck: yes, Memo: yes | OK |
| Other project net cash flow | Summary!B53 | 375.4475 | 375/375.4 | Deck: yes, Memo: yes | OK |
| Stress peak shortage | python | 244.3073 | 244/244.3 | Deck: yes, Memo: yes, DD: yes | OK |
| Downside peak shortage | python | 848.6043 | 848.6/849 | Memo: yes, DD: yes | OK |
| Ticket | python | 125.0000 | 125/125.0 | Deck: yes, Memo: yes, DD: yes | OK |
| Option B total return | python | 287.5000 | 287.5/288 | Deck: yes, Memo: yes | OK |
| Instalment | python | 23.9583 | 23.96/24/24.0 | Memo: yes | OK |

Tie-out misses: 0

Annex inputs without a source or basis note: 0

Hard-coded numeric cells on calculation sheets (excluding blue scenario parameters): 0

## 2. Arithmetic check: independent recomputation (pure-Python bisection) vs spreadsheet

| Measure | Independent | Spreadsheet | Abs difference | Result |
|---|---|---|---|---|
| Option A investor IRR | 0.292473 | 0.292473 | 1.63e-13 | OK |
| Option A investor MOIC | 2.874074 | 2.874074 | 3.55e-15 | OK |
| Option A Jalour cost of capital | 0.225574 | 0.225574 | 4.26e-14 | OK |
| Option A Jalour NPV @14% | -26.003138 | -26.003138 | 4.97e-14 | OK |
| Option B S1 investor IRR | 0.265464 | 0.265464 | 7.22e-16 | OK |
| Project NPV @14% | 307.532371 | 307.532371 | 6.82e-13 | OK |
| Project net cash flow | 595.292929 | 595.292929 | 1.14e-13 | OK |
| Peak cumulative shortage | 102.018063 | 102.018062 | 1.42e-14 | OK |
| Option B min coverage (net cash flow) | -5.394103 | -5.394103 | 8.88e-16 | OK |
| Option B quarters below 1.5x | 5.000000 | 5.000000 | 0.00e+00 | OK |
| Option B min coverage (cash available) | 3.844273 | 3.844273 | 1.33e-15 | OK |

Python engine vs spreadsheet reconciliation sheet: ALL CHECKS ZERO

## 3. Disclosure scan (every file, including document properties, notes, hidden sheets and comments)

| File | Parts scanned | Prohibited hits | Bare "250" contexts |
|---|---|---|---|
| DD_Memo.docx | 18 | 0   | 2 word/document.xml: ...00                   191,250                   106,00... | word/document.xml: ...etail                    250,000                    3... |
| Deck.pptx | 104 | 0  Allowed in context: 250 million x1 in ppt/slides/slide2.xml (units face value 2.0 x 125); 250 million x1 in ppt/slides/slide3.xml (units face value 2.0 x 125); 250 million x1 in ppt/slides/slide13.xml (units face value 2.0 x 125) | 6 ppt/slides/slide2.xml: ...with a face value of EGP 250 million (2.0x the ticket... | ppt/slides/slide3.xml: ...with a face value of EGP 250 million at launch list p... | ppt/slides/slide5.xml: ...Launch EGP 250,000 per m2, rising to 36... |
| Financial_Annex.xlsx | 22 | 0   | 10 xl/worksheets/sheet2.xml: ...32    Option_A!B5  250    B20/Inputs!$B$52  5.2... | xl/worksheets/sheet3.xml: ...137      138      139    250    140    141      142... | xl/worksheets/sheet6.xml: ...1.22527322061228       250    MIN(B12,B8)/B8  1... |
| Investment_Memo.docx | 18 | 0  Allowed in context: 250 million x3 in word/document.xml (units face value 2.0 x 125) | 15 word/document.xml: ...with a face value of EGP 250 million, twice the ticke... | word/document.xml: ...launch list (retail EGP 250,000 per m2, offices EGP... | word/document.xml: ...value                    250                       Op... |

Prohibited-term hits: 0. Bare "250" occurrences are listed for context only: they are prices or the units face value, not a programme total.

## 4. Consistency of ticket, use of funds and returns across documents

- Ticket EGP 125: Deck yes, Memo yes, DD yes, annex yes (Summary links)
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