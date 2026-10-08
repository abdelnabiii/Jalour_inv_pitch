# 00 Model Reconciliation (INTERNAL)

Source: `ALAHLY-SABBOUR MOSTAKBAL CITY -V2.xlsx` (24 sheets after the removal, 11 hidden). All figures EGP million unless stated. The model was later edited once, to remove the third project (see R10); nothing else was changed. Every number below was recomputed from the model's own cached values (`_build/extract_model.py`); the quarterly rebuild reproduces GS net 375.45 / NPV 215.31 and LA net 595.29 / NPV 307.53 exactly.

## 1. Conflicts between your prompt and the model (needs your decision)

| # | Prompt says | Model says | Verdict |
|---|---|---|---|
| C1 (resolved: 595.3 confirmed) | L'avenir outflow 862 / construction 651 / net 642 | The cash-flow engine (`Net Cash Flow -With DP`) uses outflow 908.7, construction 697.2, net **595.3**. NPV 307.5 is built on 595.3. | **Use 595.3.** `Cash Out Detail!C29 =SUM(AB29:CC29)` starts at column AB but L'avenir costs start in column V, so 46.4 of construction cost (V:AA) is dropped from `L'avenir Summary` (C18, C23) and the prompt table. This is the 595.3 vs 641.7 difference (46.4 = 908.7 - 862.3). Summary!D4 already shows 595.3. |
| C2 | Down payment 80 / 85 from Offers DP; Summary shows 200 | The "with DP" NPVs use **80 and 85**. `Summary!B3:B4` (200) are typed labels referenced only by a display cell (`C10` on each project summary); no calculation reads them. | **Present 80 / 85.** Summary!B3:B4 should be corrected before the model is shown to anyone. |
| C3 | Down payment paid in Month 3 | GS: 91.25 in M3 = 80 DP + 11.25 first instalment (confirmed). **LA: 85 DP is paid in M12**, not M3 (Net Cash Flow row 12). | LA investor money is needed by M12, not M3. |

## 2. Other inconsistencies and findings

| # | Finding | Where | Why it matters |
|---|---|---|---|
| R1 | **Delivery-payment timing (you confirmed this is contractual; kept as base, shown as a stress).** The 20-25% delivery instalment is collected in M30/33/36 (GS) and M39/42/45 (LA), but project delivery is M48 / M54. GS collects 495 in M30 when only about 26% of construction cost has been spent. Sheet delivery assumption `C5` = 36 (GS) / 48 (LA) vs Summary delivery 48 / 54. | `Green Square!N35:P36`, `C5`; `L'avenir` same | Largest model risk, now supported by contract terms (open item 5). Moving delivery payments to handover: GS peak shortage 102.5 -> **269**, NPV 215 -> 155; LA 102 -> **244**, NPV 308 -> 275. Collections slipping only 2 quarters: GS peak 236, LA 381. |
| R2 | SG&A 104.574 identical in both projects. | `Cash Out Detail` rows 40/41 | Genuine, not a copied cell: SG&A is 15% x construction, and L'avenir's construction row is **the GS stage-cost schedule shifted 6 months** (697.16 in both). LA BUA is 3.0% larger than GS, so a BUA-scaled cost would be about 21 higher. |
| R3 | Construction cost basis: stage costs 87.145 x hardcoded **8** (`Cash Out Detail` row 28/29 `=P24*8`) = 697.16, i.e. EGP 53.5k per m2 BUA. The "x8" is unlabelled. | `Stage Schedule & Costs (2)!G25`; `Cash Out Detail!P28` | Needs a label and a source (contractor quote / BOQ). Costs are flat in nominal terms for 36-54 months with no inflation. |
| R4 | **Discount rate: 14% p.a.**, labelled "Cost Of Capital" (`Green Square!C72`, used by every NPV, including L'avenir). End-of-quarter discounting, t = month/12. | `Green Square!C72`, `E73` | The CBE overnight deposit rate is **19%** (policy rates held in 2026; source below). A 14% discount rate is below the risk-free alternative investors will compare against. |
| R5 | **Sellable admin area 10,129 m2 (GS) / 10,437 m2 (LA)** vs office BUA 8,682 / 8,946 (ratio exactly 7/6). Retail uses BUA (4,341 / 4,473). | `Green Square!B7`, `L'avenir!B7` | Admin sales are priced on 16.7% more area than the BUA in the project summary. If BUA is right, GS admin sales fall by about 190 at 80% sold. Needs a sellable-area schedule. |
| R6 | **Price steps (hardcoded, sales sheets rows 3/4, EGP thousand per m2).** The price list starts at 200/125 in M3 but the first sales occur at 225/135 (GS, M9) and 250/160 (LA, M15). Launch-to-final: GS retail 225 -> 300 (+33%), office 135 -> 200 (+48%); LA retail 250 -> 360 (+44%), office 160 -> 240 (+50%). Average realised price GS 268 / 163, LA 301 / 191. You expect 40-60% appreciation: the model sits at or just below that range. | `Green Square!E3:T4`, `L'avenir!E3:T4` | Public listings for offices in Mostakbal City are about 106-120k EGP/m2 (asking, Oct 2026) vs 135 / 160 launch. Market evidence for 2025-26 is 10-16% a year; 40-60% over 3-4 years needs about 12-13% a year, the top of the evidence (see `_build/market_research.md`). Investors will challenge it. |
| R7 | **The "5.2% monthly" price escalation is not a sales-price driver.** `Price Increase` (hidden) is a *construction-cost* index (inflation x beta 1.5 + FX + progress premium: 5.2% in month 1, cumulative about 160% by Dec 2031). Nothing in the live cash flow references it; its inputs are an external link (`[25]Assumptions`). | `Price Increase!C23:BJ27` | Sales prices rise, costs do not: the model is implicitly asymmetric. A cost index of that size, if applied, would remove most of the margin. Either drop the sheet or reconcile. |
| R8 | Landlord payments: the model pays the guarantee schedule (Offers DP: GS 45/90/135/180/180/90 plus DP 80 = 800), then settles the excess in years 8-10 so that total = **35% of 100% sales value** (985.7 GS / 1,169.8 LA). Sheet text says "35% of collections whichever higher" but no max() is computed quarter by quarter. | `Net Cash Flow -With DP` rows 10/11 | The 20% retained units still carry the landlord's 35%, paid in cash by Jalour. Jalour keeps 100% of retained units and the model's net cash flow excludes their value (GS 563, LA 668 at modelled prices). |
| R9 | Units: project sheets are EGP thousand (sales/collections) and EGP (Cash Out Detail); summaries MEGP. Conversions (/1000, /1,000,000) are correct. | | Standardised to EGP million in all my files. |
| R10 | The third project that was in the original workbook, and every reference to it, has been removed from the model. Totals now sum Green Square and L'avenir only. Sheet rows below the removed rows moved up, so cell addresses below differ from the original V2 file. | whole workbook | Green Square and L'avenir figures are unchanged (checked cell by cell); only totals changed (summary NPV 522.8, net cash flow 970.7). |
| R11 | Commission 4% is charged on sales value but **paid 12 months after the sale quarter** (e.g. M9 sales are paid in M21) because `Cash Out Detail` row 34 maps sales column E to month 13. Not on collections. SG&A 15% = contingency 5% + professional fees 5% + G&A 5% of construction cost, spread with construction. No pre-construction/design cost before M13 (GS) / M19 (LA). | `Cash Out Detail` rows 33-39, 48-51 | Real design and permit spend precedes construction; I add 10.5 as an assumption (30% of the professional-fee line). |
| R12 | Sales commission, SG&A and construction are not reduced in "without DP" case; DP effect on NPV: GS -25.9, LA -21.0. | Summary rows 10-11 | Consistent. |

## 3. Headline number map

| Figure | Sheet / cell |
|---|---|
| GS net cash flow 375.45 | `Net Cash Flow -With DP!C40` (= `Summary!D3`) |
| LA net cash flow 595.29 | `Net Cash Flow -With DP!C41` (= `Summary!D4`) |
| NPV with DP 215.31 / 307.53 | `Summary!D9 / D10` (from `Net Cash Flow -With DP!C47:C48`) |
| NPV without DP 241.24 / 328.55 | `Summary!C9 / C10` |
| Peak shortage -102.5 / -102.02 | `Green Square Summary!E7`, `L'avenir Summary!E7` (`MIN` of row 49) |
| Total sales 100% / 80% | `Green Square Summary!C27 / C31`; `L'avenir Summary!C27 / C31` |
| Jalour inflow 1,267.3 / 1,504.0 | `Net Cash Flow -With DP!C16 / C17` |
| Landlord total 985.7 / 1,169.8 | `Net Cash Flow -With DP!C10 / C11` |
| Down payments 80 / 85 | `Offers DP!C9 / D9` |
| Guarantee 800 / 850 | `Offers DP!C3 / D3`, `Summary!G3:G4` |
| Costs: construction 697.16, commission 90.12 / 106.95, SG&A 104.57 | `Cash Out Detail!C28, C33/C34, C38/C39` |
| Discount rate 14% | `Green Square!C72` |

## 4. External sources used for sanity checks (public, for the reconciliation only)
- Mostakbal City office asking prices 106-120k EGP/m2: [Property Finder, Sumou Boulevard listings](https://www.propertyfinder.eg/en/plp/commercial-buy/office-space-for-sale-cairo-mostakbal-city-future-city-sumou-boulebvard-92981669.html), [Zizinia Al Mostakbal listing](https://www.propertyfinder.eg/en/plp/commercial-buy/office-space-for-sale-cairo-new-cairo-city-the-5th-settlement-mostakbal-city-compounds-zizinia-al-mostakbal-94906776.html). Listing asks, not transactions.
- CBE overnight deposit rate 19.0%, held at the July 2026 MPC: [Egypt Today](https://www.egypttoday.com/Article/3/148261/CBE-keeps-interest-rates-unchanged-for-third-consecutive-meeting-in), [Trading Economics](https://tradingeconomics.com/egypt/interest-rate).
