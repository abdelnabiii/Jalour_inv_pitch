# QA report: Green Square and L'avenir combined pack, ticket EGP 125m (62.5 + 62.5)

## 1. Number tie-out (annex vs deck, memo, DD memo)

| Figure | Annex cell | Value | Found in | Result |
|---|---|---|---|---|
| Combined ticket | Combined_Summary!B18 | 125.0000 | Deck: yes, Memo: yes, DD: yes | OK |
| Option A face value (combined) | Combined_Summary!B21 | 250.0000 | Deck: yes, Memo: yes | OK |
| Option A value at delivery list | Combined_Summary!B22 | 364.3967 | Memo: yes | OK |
| Option A investor IRR base | Combined_Summary!B23 | 0.2768 | Deck: yes, Memo: yes | OK |
| Option A MOIC base | Combined_Summary!B24 | 2.8277 | Deck: yes | OK |
| Option A IRR downside | Combined_Summary!B25 | 0.1946 | Deck: yes, Memo: yes | OK |
| Option A IRR upside | Combined_Summary!B27 | 0.3057 | Deck: yes, Memo: yes | OK |
| Jalour cost of capital, Option A | Combined_Summary!B29 | 0.2156 | Memo: yes | OK |
| Option B total return | Combined_Summary!B31 | 287.5000 | Deck: yes, Memo: yes | OK |
| Option B investor IRR | Combined_Summary!B32 | 0.2655 | Deck: yes, Memo: yes | OK |
| Option B MOIC | Combined_Summary!B33 | 2.3000 | Deck: yes, Memo: yes | OK |
| Combined lowest position, base | Combined_Position!B11 | -167.9546 | Deck: yes, Memo: yes, DD: yes | OK |
| GS base peak shortage | python | 102.5000 | Deck: yes, Memo: yes | OK |
| LA base peak shortage | python | 102.0181 | Deck: yes, Memo: yes | OK |
| GS net cash flow | python | 375.4475 | Deck: yes, Memo: yes | OK |
| LA net cash flow | python | 595.2929 | Deck: yes, Memo: yes | OK |
| GS NPV | python | 215.3114 | Deck: yes, Memo: yes | OK |
| LA NPV | python | 307.5324 | Deck: yes, Memo: yes | OK |
| Combined peak, handover payments | python | 318.3333 | Deck: yes, Memo: yes, DD: yes | OK |
| Combined peak, downside | python | 1088.9735 | Deck: yes, Memo: yes, DD: yes | OK |
| Combined low month | python | 15.0000 | Deck: yes, Memo: yes | OK |
| Combined positive from month | python | 30.0000 | Deck: yes, Memo: yes | OK |
| Instalment per project | python | 11.9792 | Memo: yes | OK |
| Option B per project | python | 143.7500 | Deck: yes, Memo: yes | OK |

Tie-out misses: 0

## 2. Arithmetic check: independent pure-Python recomputation vs spreadsheet

| Measure | Independent | Spreadsheet / engine | Abs difference | Result |
|---|---|---|---|---|
| Option A combined investor IRR | 0.276787 | 0.276787 | 1.39e-15 | OK |
| Option A combined MOIC | 2.827718 | 2.827718 | 2.66e-15 | OK |
| Option A Jalour cost of capital | 0.215630 | 0.215630 | 2.75e-15 | OK |
| Option A Jalour NPV @14% | -24.934707 | -24.934707 | 1.03e-13 | OK |
| Option B combined investor IRR | 0.265464 | 0.265464 | 6.11e-16 | OK |
| GS project NPV @14% | 215.311393 | 215.311393 | 1.71e-13 | OK |
| GS net cash flow | 375.447528 | 375.447528 | 1.71e-13 | OK |
| LA project NPV @14% | 307.532371 | 307.532371 | 5.12e-13 | OK |
| LA net cash flow | 595.292929 | 595.292929 | 3.41e-13 | OK |
| Combined lowest cumulative position | -167.954631 | -167.954631 | 0.00e+00 | OK |

Annex reconciliation sheet final check: ALL CHECKS ZERO

## 3. Disclosure scan (every part of every file)

Both project names are permitted in this pack: it is the single combined investor pack.

| File | Parts scanned | Prohibited hits | Confidentiality statement |
|---|---|---|---|
| DD_Memo.docx | 18 | 0  | yes |
| Deck.pptx | 106 | 0  | yes |
| Financial_Annex.xlsx | 34 | 0  | yes |
| Investment_Memo.docx | 18 | 0  | yes |
| NDA_Template.docx | 18 | 0  | yes |
- Neutral funding wording in Deck: yes
- Neutral funding wording in Memo: yes
- Neutral funding wording in DD: yes
- Further capital clause in Deck: yes
- Further capital clause in Memo: yes
- Further capital clause in DD: yes
- No cross-collateral in Deck: yes
- No cross-collateral in Memo: yes
- No cross-collateral in DD: yes
- Confidential banner in Deck: yes
- Confidential banner in Memo: yes
- Confidential banner in DD: yes
- Confidential banner in NDA: yes
- NDA covers all information in NDA: yes

Disclosure and wording failures: 0

- NDA template, Green Square named: present
- NDA template, L'avenir named: present
- NDA template, Covers financial information of both projects: present
- NDA template, Non-circumvention: present
- NDA template, Return or destroy: present
- NDA template, Term: present
- NDA template, Governing law: present
- NDA template, No obligation / no exclusivity: present

**Total failures requiring attention: 0**