import sys, json, os; sys.path.insert(0, '.')
from docx_helpers import *
import packdata as PD, figs
MK = json.load(open('market.json')); MF = {f['id']: f for f in MK['facts']}
BAL = 'The balance of project funding is provided by Jalour sponsor equity, project collections and other capital sources.'
def build_memo(k, T, path, workdir):
    D = PD.build(k, T); I = D['info']; A = D['A']; B = D['B']; F = D['fund']; S1 = B['S']['S1']['res']; S2 = B['S']['S2']['res']; S3 = B['S']['S3']['res']
    NAME = I['name']; figp = figs.make(D, os.path.join(workdir, 'figs'))
    inst = B['mult']*T/12; pm = B['pay_months']; t0 = 3 if k == 'GS' else 12
    pr = D['prices']; pl_c, pl_a = pr['launch']; pf_c, pf_a = pr['final']; ap_c, ap_a = pr['avg']
    g_c = pf_c/pl_c-1; g_a = pf_a/pl_a-1
    net = np_net = D['series']['net']; months = D['series']['months']
    neg_m = [m for m in B['pay_months'] if D['series']['net'][m//3-1] < 0]; peak_txt = f'months {neg_m[0]} to {neg_m[-1]}' if neg_m else 'none'
    def one(toc=None):
        d = Doc(f'{NAME} - Investment Memorandum (EGP {T} million)', f'Jalour Developments  |  {NAME} Investment Memorandum  |  Strictly private and confidential'); dd = d.d
        # ---------- cover ----------
        for _ in range(5): dd.add_paragraph()
        d.p('STRICTLY PRIVATE AND CONFIDENTIAL', bold=True, size=10, color=BRASS)
        d.p(NAME, bold=True, size=34, color=NAVY, after=2)
        d.p('Commercial and office development, El Mostakbal City, Cairo', size=15, color=INK, after=14)
        d.p('INVESTMENT MEMORANDUM', bold=True, size=14, color=NAVY, after=2)
        d.p(f'Private placement of EGP {T} million (USD {T/48:,.1f} million at EGP 48 per USD)', size=12, after=2)
        d.p('Option A: units. Option B: cash.', size=12, after=36)
        d.p('Issued by Jalour Developments, a subsidiary of Al Jalal Holding, Cairo', size=10.5, after=2)
        d.p('October 2026. Draft for discussion with the named investor only.', size=10.5, color=GREY)
        # ---------- notice ----------
        d.h1('Important notice')
        for t in ['This memorandum has been prepared by Jalour Developments (Jalour) for the sole use of the investor to whom it is addressed. It is confidential. It must not be copied, distributed or disclosed to any other person without Jalour\'s written consent.',
                  'It is not an offer to the public and is not a prospectus. It has not been reviewed or approved by any regulator. Whether this placement requires any regulatory filing, approval or exemption in Egypt or elsewhere is a matter for legal counsel; nothing here is legal, tax or investment advice.',
                  'All forecasts, returns, multiples, internal rates of return, cash flows and market statements are projections or reported third-party statements, not statements of fact or guarantees. Actual results may differ materially. Every return figure in this memorandum is labelled with its case (base, downside or upside).',
                  'Figures come from the Jalour financial model for the project and are reproduced in the Financial Annex, whose Summary sheet is the reference for every number below. Market data are taken from third-party publications identified in the text and in Appendix D; Jalour has not independently verified them. Property listings are asking prices, not transaction prices.',
                  'Items shown in square brackets are pending from Jalour. Terms are indicative until set out in definitive agreements, which prevail over this memorandum.',
                  'Jalour may raise further capital at project or holding level, subject to the investor\'s stated rights as to ranking, collateral and information. This memorandum creates no exclusivity in favour of any investor.',
                  'The investment has no exposure to other Jalour projects: the project is held in its own company, with no cross-collateral and no cross-default to any other project.']: d.p(t)
        d.h1('Contents')
        toc_items = ['1. Executive summary', '2. Transaction overview', '3. Investment thesis', '4. Market analysis', '5. Project description', '6. The Al Ahly Sabbour agreement and consents', '7. Sponsor and management', '8. Business plan and assumptions', '9. Sales and collections model', '10. Cost model',
                     '11. Funding plan and use of funds', '12. Investment terms', '13. Returns analysis and sensitivities', '14. Security and ranking', '15. Governance and information rights', '16. Exit and liquidity', '17. Risk factors', '18. Tax and legal considerations', 'Appendices']
        for t in toc_items:
            par = dd.add_paragraph(); par.paragraph_format.space_after = Pt(3)
            par.paragraph_format.tab_stops.add_tab_stop(Cm(16.3), alignment=2, leader=1)
            r = par.add_run(t + '\t' + (str(toc.get(t, '')) if toc else '')); r.font.size = Pt(11)
        # ---------- 1 Executive summary ----------
        d.h1('1. Executive summary')
        d.p(f'Jalour is raising EGP {T} million from a single investor to part-fund {NAME}, a {fmt(I["bua"],0)} m2 commercial and office development of {I["height"]} buildings on a {fmt(I["land"],0)} m2 plot in El Mostakbal City, east of New Cairo. The land belongs to Al Ahly Sabbour. Jalour is developer and operator: it designs, obtains permits, builds to core-and-shell standard, markets, sells, collects and operates. Collections are shared 35% to the landowner (subject to a minimum guarantee of EGP {fmt(I["guarantee"],0)} million over six years) and 65% to Jalour.')
        d.p('The investor chooses one of two instruments at signing:')
        d.bullets([f'**Option A, units.** Units with a face value of EGP {fmt(A["face"],0)} million, twice the ticket, priced at the launch list (retail EGP {fmt(pl_c*1000,0)} per m2, offices EGP {fmt(pl_a*1000,0)} per m2), allocated from the first two sales tranches. At the model\'s final list prices the units are worth EGP {fmt(A["value_at_handover"],0)} million. Projected investor IRR: {pct(A["inv"]["Base"]["irr"])} base, {pct(A["inv"]["Downside"]["irr"])} downside, {pct(A["inv"]["Upside"]["irr"])} upside.',
                   f'**Option B, cash.** EGP {fmt(B["mult"]*T,0)} million ({mx(B["mult"],1)} the ticket) in 12 equal quarterly instalments of EGP {fmt(inst,2)} million after a 24-month grace period, from month {pm[0]} to month {pm[-1]}. Projected investor IRR {pct(S1["Base"]["irr"])}, contractual if Jalour pays.'])
        d.table([['Key figures (EGP million unless stated)', 'Value'],
                 ['Investor ticket', fmt(T, 1)], ['Modelled peak cumulative cash shortage (month ' + str(D['base']['peak_m']) + ')', fmt(F['model_peak'], 1)], ['Total uses at the funding peak', fmt(F['total'], 1)],
                 ['Total sales at 80% of units (sales plan)', fmt(I['total80'], 0)], ['Jalour net cash flow, nominal', fmt(D['base']['net'], 1)], ['NPV of Jalour net cash flow at 14%', fmt(D['base']['npv'], 1)],
                 ['Handover (project month)', str(I['delivery'])], ['Option A units, face value', fmt(A['face'], 0)], ['Option B total return', fmt(B['mult']*T, 0)]], [11.0, 5.3], 'Headline figures', 1)
        d.p('Three points deserve the investor\'s attention before the detail.', bold=True, keep=True)
        d.bullets([f'**Pricing.** The plan relies on list prices rising from EGP {fmt(pl_c*1000,0)} to {fmt(pf_c*1000,0)} per m2 for retail (+{g_c*100:.0f}%) and from {fmt(pl_a*1000,0)} to {fmt(pf_a*1000,0)} for offices (+{g_a*100:.0f}%) across the sales window. Recent evidence of price growth is lower than this at its low end and similar at its high end (section 4). Option B removes unit price exposure; Option A does not.',
                   f'**Timing of cash.** The model collects the delivery payment (20% to 25% of price) on a contractual milestone ahead of handover. If it were collected at handover the peak funding need would be EGP {fmt(F["stress_peak"],0)} million instead of EGP {fmt(F["model_peak"],1)} million. Jalour has said it will fund project funding gaps above project cash (section 11).',
                   f'**Payout capacity (Option B).** Net project cash flow is negative in the construction peak ({peak_txt}, inside the payout window), so quarter-by-quarter coverage of instalments is below 1.5x in {S1["Base"]["strict_below"]} of {S1["Base"]["n"]} quarters. Pooled cash including the ticket covers every instalment at least {mx(S1["Base"]["liq_min"],1)} in the base case (section 13).'])
        d.p(BAL, italic=True)
        # ---------- 2 Transaction overview ----------
        d.h1('2. Transaction overview')
        d.h2('2.1 Parties and instrument')
        d.table([['Item', 'Description'], ['Sponsor and developer', 'Jalour Developments, the real estate arm of Al Jalal Holding, Cairo'], ['Landowner', 'Al Ahly Sabbour'], ['Issuer', 'The project company for this development [name and jurisdiction to be confirmed]'],
                 ['Investor', '[Name of investor]'], ['Instrument', 'Contractual entitlement to units (Option A) or to cash payments (Option B). Whether it is documented as equity, a loan or a contractual entitlement is to be confirmed by counsel and affects tax, regulation and enforceability'], ['Ticket', f'EGP {fmt(T,0)} million'],
                 ['Choice', 'Option A or Option B, selected at signing; not both'], ['Balance of funding', BAL]], [4.2, 12.1], 'Principal terms', -1, bold_first_col=True)
        d.h2('2.2 Conditions and timetable')
        d.bullets(['Due diligence by the investor and its advisers on the data room (Due Diligence Memorandum).', 'Written consent of Al Ahly Sabbour to the unit allocation (Option A) and to any assignment or encumbrance that affects its share.', 'Definitive agreements, including security documents reviewed by counsel for both sides.', f'Funding in cash into the project company, expected for project month {t0} (the month the landlord down payment falls due in the model).'])
        d.h2('2.3 No exclusivity; further capital')
        d.p('Jalour may raise further capital at project or holding level, subject to the investor\'s stated rights: ranking of the investor\'s claim, collateral over the earmarked units, and information rights. The project company will not grant security over the investor\'s earmarked units or rank any new instrument ahead of the investor\'s entitlement without the consent given in the definitive agreements.')
        d.h2('2.4 Separation from other projects')
        d.p('The project is held in its own company. There is no cross-collateral and no cross-default with any other Jalour project, and the investor has no exposure to other Jalour projects. The funds are project funding, not a restricted or segregated account; payment protections are contractual and described in section 14.')
        # ---------- 3 Thesis ----------
        d.h1('3. Investment thesis', page_break=False)
        d.p('The case for the investment, with the main reservation against each point:', keep=True)
        d.table([['Thesis', 'Reservation'],
                 ['A defined, permitted-scale asset: two G+2 buildings of ' + fmt(I['bua'], 0) + ' m2 in a state-planned city of about 11,000 feddans and about 1 million planned residents.', 'Mostakbal City is early in its build-out; current resident population and occupancy were not found in public sources.'],
                 ['Land is owned by an established developer, Al Ahly Sabbour, and is not an acquisition cost for Jalour: the landowner is paid from collections.', 'The landlord\'s guarantee (EGP ' + fmt(I['guarantee'], 0) + ' million over six years) is payable even if sales are slow.'],
                 [f'Sales begin at a launch list of EGP {fmt(pl_c*1000,0)} (retail) and {fmt(pl_a*1000,0)} (offices) per m2, with list price steps built into the plan.', 'Asking prices for Mostakbal City offices today are about EGP 106,000 to 120,000 per m2; the plan assumes higher.'],
                 ['Project funding peaks at EGP ' + fmt(F['model_peak'], 1) + ' million in the base case and the project generates EGP ' + fmt(D['base']['net'], 0) + ' million of net cash flow before the value of retained units.', 'Funding peak is EGP ' + fmt(F['stress_peak'], 0) + ' million if delivery payments arrive at handover.'],
                 ['Two ways to participate: units with price upside, or cash with a fixed return.', 'Option A carries price and resale risk; Option B depends on Jalour\'s ability to fund payouts during the construction peak.']], [8.3, 8.0], 'Thesis and reservations', -1, font=9.5)
        d.p('Jalour retains 20% of the units. They sit outside the sales plan, outside the cash flows shown here, and outside this offering.')
        # ---------- 4 Market ----------
        d.h1('4. Market analysis')
        d.p('Market statements below are reported by the third parties named. They are drawn from published summaries as at October 2026; Jalour has not audited them. Where sources use different definitions (for example, Knight Frank and JLL measure Cairo office stock differently) the figures are not additive.')
        d.h2('4.1 Cairo and New Cairo offices')
        d.bullets([MF['kf_stock']['text'] + f' ({MF["kf_stock"]["src"]}).', MF['kf_nc']['text'] + f' ({MF["kf_nc"]["src"]}).', MF['jll_vac']['text'] + f' ({MF["jll_vac"]["src"]}).', MF['kf_rent']['text'] + f' ({MF["kf_rent"]["src"]}).'])
        d.p('Reading across: occupancy and rents are firm today, but the pipeline is large relative to stock. New Cairo is where most of that supply will land, so absorption of new supply by 2029 is the main market risk for the offices.')
        d.h2('4.2 Retail and commercial')
        d.bullets([MF['jll_retail']['text'] + f' ({MF["jll_retail"]["src"]}).', MF['hap']['text'] + f' ({MF["hap"]["src"]}).', 'Retail rents rose 0.8% to 3.6% y/y in Q1 2026 depending on format (JLL via Invest-Gate); vacancy fell from 9.2% in Q2 2024 to 7.2% in Q2 2025 (JLL).'])
        d.p('The product here, ground-floor retail beneath two office floors in a G+2 form, is close to the mixed-use commercial boulevards now being launched in Mostakbal City. Retail sales in a new city depend on resident and worker footfall that builds over years.')
        d.h2('4.3 Mostakbal City')
        d.bullets([MF['mc_scale']['text'] + f' ({MF["mc_scale"]["src"]}).', MF['mc_infra']['text'] + f' ({MF["mc_infra"]["src"]}).', 'Al Ahly Sabbour is active in the city (residential schemes and a mixed-use plot), which supports the landowner\'s interest in the success of this development (Invest-Gate).'])
        d.h2('4.4 Pricing evidence and the plan')
        d.table([['Measure (EGP per m2)', 'Evidence (asking prices)', 'Plan: launch', 'Plan: final list', 'Plan: average'],
                 ['Offices', '106,000 to 120,000 (Mostakbal City listings)', fmt(pl_a*1000, 0), fmt(pf_a*1000, 0), fmt(ap_a*1000, 0)],
                 ['Retail', '139,000 to 325,000 (New Cairo listings, individual)', fmt(pl_c*1000, 0), fmt(pf_c*1000, 0), fmt(ap_c*1000, 0)]], [3.0, 5.8, 2.5, 2.5, 2.5], 'Plan prices against market asking prices', 2, font=9)
        d.p(f'The plan\'s launch price for offices ({fmt(pl_a*1000,0)}) is above current Mostakbal City asking prices, and it then rises by {g_a*100:.0f}%. New Cairo secondary-market prices rose 3.1% y/y in Q1 2026 (JLL via Enterprise); broker estimates for 2026 price growth are 10% to 16% a year (a blog estimate, not an index). A 40% to 60% rise over three to four years needs about 12% to 13% a year, which is at the upper end of that range and depends on continued inflation (14.5% in August 2026). Sales prices may therefore be lower or slower than planned; section 13 shows the effect.')
        d.h2('4.5 Demand drivers and macro')
        d.bullets([MF['cbe']['text'] + f' ({MF["cbe"]["src"]}).', MF['dev_h1']['text'] + f' ({MF["dev_h1"]["src"]}).', MF['payplan']['text'] + f' ({MF["payplan"]["src"]}).', 'Real estate is widely used in Egypt as an inflation hedge, but affordability is stretched and the share of investor buyers fell in 2025 (Enterprise).'])
        d.h2('4.6 Regulation of off-plan sales')
        d.p(MF['decree']['text'] + f' ({MF["decree"]["src"]}). Counsel should confirm which rules apply to these plots and to a private placement of this kind, including any requirement for a separate project account. Resale of contracts is common but developer approval and fees are contractual; no specific rule on developer consent to resale was found.')
        d.figure(figp['sales'], 'Sales value by quarter of sale', 'Jalour financial model; Financial Annex, Source_Data sheet')
        # ---------- 5 Project ----------
        d.h1('5. Project description')
        d.table([['Item', 'Detail'], ['Land area', f'{fmt(I["land"],0)} m2'], ['Footprint (30% of land)', f'{fmt(I["footprint"],0)} m2'], ['Height', I['height']], ['Retail built-up area', f'{fmt(I["retail"],0)} m2'], ['Office built-up area', f'{fmt(I["office"],0)} m2'], ['Total built-up area', f'{fmt(I["bua"],0)} m2'],
                 ['Finishing standard', 'Core and shell'], ['Sales start', f'Project month {I["sales_start"]}'], ['Construction', f'Months {I["con_start"]} to {I["delivery"]} ({I["con_months"]} months)'], ['Handover', f'Project month {I["delivery"]}'], ['Retained by Jalour', '20% of units (not offered, not in the cash flows)']], [6.0, 10.3], 'Project facts (source: Jalour financial model)', -1, bold_first_col=True)
        d.p(f'The development comprises retail on the ground floor and offices above, in {I["height"]} buildings over a 30% footprint. Retail accounts for {fmt(I["retail"]/I["bua"]*100,0)}% of built-up area and offices for {fmt(I["office"]/I["bua"]*100,0)}%. Jalour is responsible for design, permits and licences, construction to core and shell, project branding, marketing, sales and collection, and for operation, management and maintenance directly or through a facility management company.')
        d.p(f'The model sells offices on a sellable area of {fmt(I["admin_sales_area"],0)} m2, which is larger than the {fmt(I["office"],0)} m2 of office built-up area. Jalour is to provide the area schedule that reconciles the two (open item); the investor should require this in due diligence.')
        d.p('Phasing: sales launch in project month ' + str(I['sales_start']) + ', construction starts in month ' + str(I['con_start']) + ' and runs ' + str(I['con_months']) + ' months to handover in month ' + str(I['delivery']) + '. Permits, licences and title status are listed in the Due Diligence Memorandum [status pending from Jalour].')
        # ---------- 6 Landlord ----------
        d.h1('6. The Al Ahly Sabbour agreement and consents')
        d.p('The following terms are taken from the Jalour financial model and are to be verified against the signed agreement in the data room.')
        d.table([['Term', 'Detail'], ['Landlord share of collections', '35%'], ['Developer (Jalour) share', '65%'], ['Minimum guaranteed amount', f'EGP {fmt(I["guarantee"],0)} million over six years'], ['Down payment', f'EGP {fmt(I["dp"],0)} million (10% of the guarantee), paid up front'],
                 ['Payment rule', '35% of collections or the guaranteed instalment, whichever is higher'], ['Settlement of any excess over the guarantee', f'From year {I["excess_years"]}'], ['Jalour scope', 'Design, permits, construction (core and shell), marketing, sales, collection, project name, operation and maintenance'], ['Delivery', f'Month {I["delivery"]}']], [6.0, 10.3], 'Agreement terms', -1, bold_first_col=True)
        sch = [I['dp']] + I['sched']
        d.table([['EGP million', 'Down payment', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5', 'Year 6', 'Total'], ['Guaranteed payment'] + [fmt(x, 0) for x in sch] + [fmt(sum(sch), 0)]], [3.4, 1.9, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.6], 'Minimum guarantee schedule', 1, font=9)
        d.p(f'In the plan the landlord receives EGP {fmt(D["base"]["land"],0)} million in total, equal to 35% of the value of all units at modelled prices, including the 20% retained by Jalour. EGP {fmt(D["base"]["land"]-I["guarantee"],0)} million of that is the excess over the guarantee, settled from year {I["excess_years"]}. Jalour pays the landlord share on the retained units in cash and keeps those units.')
        d.h2('6.1 Consents')
        d.p('Al Ahly Sabbour\'s written consent is required for:', keep=True)
        d.bullets(['allocation of units to the investor (Option A) and the investor\'s right to assign the contract;', 'any assignment of collections or encumbrance on units that affects the landlord\'s 35% share;', 'earmarking of unsold units as collateral (Option B);', 'any variation of the unit mix, prices or payment plans that changes the landlord\'s entitlement.'])
        d.p('Under the structure proposed here Jalour bears the landlord\'s 35% on the units allocated to the investor: landlord payments are unchanged and the landlord does not waive any share. Consent status: [to be provided by Jalour]. Receipt of consent is a condition precedent to signing.')
        # ---------- 7 Sponsor ----------
        d.h1('7. Sponsor and management', page_break=False)
        d.p('Jalour Developments is the real estate arm of Al Jalal Holding, Cairo. In this project it acts as developer and operator and carries all design, permit, construction, marketing, sales, collection and operating costs.')
        d.table([['Information', 'Status'], ['Completed projects, area delivered, units handed over', '[To be provided by Jalour]'], ['Management team biographies', '[To be provided by Jalour]'], ['Audited financial statements, Jalour and Al Jalal Holding', '[To be provided by Jalour]'], ['References from landlords, contractors and buyers', '[To be provided by Jalour]'], ['Contractor and consultant appointments for this project', '[To be provided by Jalour]']], [10.0, 6.3], 'Sponsor information pending', -1)
        d.p('Jalour intends to fund project funding gaps, including those caused by timing differences in collections, from sponsor equity or other capital sources it procures. The form, amount and conditions of that undertaking are to be documented. It is a commitment to fund the project, not a guarantee of the investor\'s return.')
        # ---------- 8 Business plan ----------
        d.h1('8. Business plan and assumptions')
        d.table([['Assumption', 'Value', 'Basis'], ['Discount rate', '14% a year', 'Jalour model (labelled cost of capital)'], ['Sales plan', '80% of units sold; 20% retained', 'Jalour model'], ['Sales window', f'Months {I["sales_start"]} to {D["tranches"][-1]["month"]}', 'Jalour model'], ['Launch list prices', f'Retail {fmt(pl_c*1000,0)}; offices {fmt(pl_a*1000,0)} EGP per m2', 'Jalour model, first sales column'],
                 ['Final list prices', f'Retail {fmt(pf_c*1000,0)}; offices {fmt(pf_a*1000,0)} EGP per m2', 'Jalour model'], ['Payment plans', 'Down payment 5% to 20%; equal second payment after 3 months; instalments 2 to 8 years; delivery payment 20% to 25%', 'Jalour model'], ['Construction cost', f'EGP {fmt(D["base"]["cons"],1)} million, flat in nominal terms', 'Jalour model'],
                 ['Sales commission', '4% of sales value, paid 12 months after the sale quarter', 'Jalour model'], ['SG&A and others', '15% of construction cost (5% contingency, 5% professional fees, 5% overhead)', 'Jalour model'], ['Landlord', '35% share; guarantee schedule; excess from year ' + str(I['excess_years']), 'Agreement terms (section 6)'],
                 ['Timeline', f'Month 1 is the start of the model; investor funds in month {t0}', 'Assumption']], [3.6, 7.4, 5.3], 'Key assumptions', -1, font=9)
        d.p('The model is quarterly, in EGP million. Net cash flow equals collections less landlord payments, construction, commission and SG&A. It excludes the retained units, any financing costs and taxes, and any cost escalation.')
        # ---------- 9 Sales ----------
        d.h1('9. Sales and collections model')
        rows = [['Quarter of sale', 'Retail (EGP m)', 'Offices (EGP m)', 'Retail price (EGP k/m2)', 'Office price (EGP k/m2)']]
        for t in D['tranches']: rows.append([f'Month {t["month"]}', fmt(t['comm'], 1), fmt(t['admin'], 1), fmt(t['pc'], 0), fmt(t['pa'], 0)])
        rows.append(['Total', fmt(sum(t['comm'] for t in D['tranches']), 1), fmt(sum(t['admin'] for t in D['tranches']), 1), '', ''])
        d.table(rows, [3.4, 3.0, 3.2, 3.4, 3.3], 'Sales plan by tranche (80% of units)', 1, font=9, total_row=True)
        d.p(f'Total sales at 100% of units are EGP {fmt(I["total100"],0)} million at modelled prices; the plan sells 80%, EGP {fmt(I["total80"],0)} million. The first two tranches (months {D["A"]["launch_months"][0]} and {D["A"]["launch_months"][1]}) carry the launch list and total EGP {fmt(A["launch_value"],0)} million; Option A units come from these.')
        ann = D['annual']
        d.table([['Project year', 'Collections', 'Landlord payments', 'Construction, commission, SG&A', 'Net cash flow', 'Cumulative']] + [[f'Year {r["year"]}', fmt(r['coll'], 1), fmt(-r['land'], 1), fmt(-r['cost'], 1), fmt(r['net'], 1), fmt(r['cum'], 1)] for r in ann[:11]] + [['Total', fmt(sum(r['coll'] for r in ann), 1), fmt(-sum(r['land'] for r in ann), 1), fmt(-sum(r['cost'] for r in ann), 1), fmt(sum(r['net'] for r in ann), 1), '']],
                [2.4, 2.5, 2.9, 3.6, 2.5, 2.4], 'Annual cash flow, EGP million (years 1 to 11; year 1 = months 1 to 12)', 1, font=9, total_row=True)
        d.figure(figp['cash'], 'Project cash flow by year', 'Jalour financial model; Financial Annex, Project_CF sheet')
        d.p('The model collects the delivery payment on a contractual milestone ahead of handover, which produces a large collection in project month ' + str(30 if k == 'GS' else 45) + '. This timing is a key sensitivity (sections 11 and 13). Evidence of the contractual milestone is to be provided in the data room.')
        # ---------- 10 Cost ----------
        d.h1('10. Cost model', page_break=False)
        d.table([['Cost line', 'EGP million', 'Basis'], ['Construction (core and shell)', fmt(D['base']['cons'], 1), f'About EGP {fmt(D["base"]["cons"]*1e6/I["bua"]/1000,1)} thousand per m2 of built-up area; stage schedule over 36 months'], ['Sales commission', fmt(D['base']['commission'], 1), '4% of sales value, paid 12 months after each sale quarter'], ['SG&A and others', fmt(D['base']['sga'], 1), '15% of construction: contingency 5%, professional fees 5%, overhead 5%'], ['Total', fmt(D['base']['cost'], 1), '']], [5.0, 2.6, 8.7], 'Cost model', 1, font=9, total_row=True)
        d.bullets(['Cost basis (contractor pricing, bill of quantities) is to be provided in the data room; the model uses a flat nominal cost with no inflation over the build period.', f'A 10% overrun lowers project NPV from EGP {fmt(D["base"]["npv"],0)} million to EGP {fmt(D["sens"]["cost"][1.1]["npv"],0)} million; a 20% overrun to EGP {fmt(D["sens"]["cost"][1.2]["npv"],0)} million.', 'Pre-construction costs (design, permits) are not separately modelled; section 11 adds an allowance of EGP ' + fmt(F['uses']['precon'], 1) + ' million as an assumption.'])
        # ---------- 11 Funding ----------
        d.h1('11. Funding plan and use of funds')
        u = F['uses']
        d.table([['Uses at the funding peak', 'EGP million'], ['Down payment to the landowner', fmt(u['dp'], 1)], ['Guarantee instalments and working capital to the peak, net of collections', fmt(u['guar_wc'], 1)], ['Design, permits and pre-construction (assumption)', fmt(u['precon'], 1)], ['Total uses', fmt(F['total'], 1)]], [12.0, 4.3], 'Uses', 1, total_row=True)
        d.table([['Sources', 'EGP million', '% of uses'], ['Investor ticket' + (' (applied to uses)' if F['excess_over_uses'] > 0 else ''), fmt(F['applied'], 1), pct(F['share_uses'], 0)], ['Balance of project funding', fmt(F['balance'], 1), pct(F['balance']/F['total'], 0)], ['Ticket above modelled uses (general project liquidity)', fmt(F['excess_over_uses'], 1), ''], ['Total', fmt(F['applied']+F['balance']+F['excess_over_uses'], 1), '']], [9.0, 3.6, 3.7], 'Sources', 1, total_row=True)
        d.p(BAL, italic=True)
        d.p(f'The ticket equals {pct(F["share_peak"],0)} of the modelled peak cumulative shortage of EGP {fmt(F["model_peak"],1)} million (month {D["base"]["peak_m"]}).')
        if F['excess_over_peak'] > 0:
            d.callout(f'**Ticket above the modelled peak.** The ticket exceeds the modelled peak shortage by EGP {fmt(F["excess_over_peak"],1)} million ({fmt(F["excess_over_uses"],1)} million above total modelled uses). The excess is general project liquidity, not held in a restricted account. It is available for the timing risk on delivery payments (peak shortage EGP {fmt(F["stress_peak"],0)} million if they arrive at handover) and for pre-construction cost overruns. Returns are contractual on the full ticket, so Jalour\'s cost of capital is measured on the full amount.')
        d.h2('11.1 Funding gaps in stress cases')
        d.table([['Case', 'Peak cumulative shortage (EGP m)', 'NPV (EGP m)']] + [[nm, fmt(-v['peak'], 0), fmt(v['npv'], 0)] for nm, v in D['proj'].items()], [8.8, 4.0, 3.5], 'Project cash position by case, before any investor instrument', 1, font=9)
        d.p('Jalour will fund the gap between project cash and the cases above from sponsor equity or other capital sources it procures; it does not expect the investor to fund it. The investor should note that in the downside and combined cases the funding need is several times the ticket and depends on Jalour\'s capacity to provide it.')
        # ---------- 12 Terms ----------
        d.h1('12. Investment terms')
        d.h2('12.1 Option A: units')
        d.table([['Term', 'Detail'], ['Face value', f'EGP {fmt(A["face"],0)} million: {A["face_x"]:.1f}x the ticket, priced at the launch list'], ['Source of units', f'Unsold units in the first two sales tranches (months {A["launch_months"][0]} and {A["launch_months"][1]}), EGP {fmt(A["launch_value"],0)} million at list; the investor receives {pct(A["f"],1)} of them. The 20% of units retained by Jalour are not used'],
                 ['Unit mix', f'Pro rata to those tranches: retail EGP {fmt(A["retail_face"],1)} million ({fmt(A["area_retail"],0)} m2) and offices EGP {fmt(A["office_face"],1)} million ({fmt(A["area_office"],0)} m2)'], ['Value at final list', f'EGP {fmt(A["value_at_handover"],0)} million ({A["appr"]:.3f}x face), before resale costs'],
                 ['Day-one holding', 'A registered unit allocation and reservation contract, assignable with the consents below'], ['Landlord share', f'Jalour bears the landlord 35% on the allocated units (EGP {fmt(A["jal"]["landlord_on_units"],0)} million at face); landlord payments are unchanged'], ['Resale (assumption)', f'Handover in month {I["delivery"]}; resale over four quarters from month {I["delivery"]+3}; 3% cost (broker and developer transfer fee)'],
                 ['Consents', 'Jalour consent and landlord acknowledgement for assignment of a contract before handover [fee to be agreed]']], [3.8, 12.5], 'Option A terms', -1, bold_first_col=True, font=9.5)
        d.p('The investor holds no physical unit until handover. Before handover the investor can sell only the contract, with consent, and a buyer is likely to require a discount. Section 16 shows the effect of early assignment. Jalour cannot promise liquidity that the structure does not provide.')
        d.h2('12.2 Option B: cash')
        d.table([['Term', 'Detail'], ['Total return', f'EGP {fmt(B["mult"]*T,0)} million ({mx(B["mult"],1)} the ticket), including return of capital'], ['Grace period', '24 months from funding, no payments'], ['Payments', f'12 equal quarterly instalments of EGP {fmt(inst,2)} million in months {pm[0]} to {pm[-1]}'], ['Obligor', 'The project company; no guarantee of return by Jalour or Al Jalal Holding'], ['Default', 'Instalment unpaid for 30 days: payment priority covenant applies; step-in right over earmarked units [terms to be agreed]'], ['Funding gaps', 'Jalour will fund project funding gaps; this supports delivery of the project and is not a guarantee of the return']], [3.8, 12.5], 'Option B terms', -1, bold_first_col=True, font=9.5)
        sched = [['Instalment', 'Project month', 'Amount (EGP m)', 'Cumulative (EGP m)']] + [[str(i+1), str(m), fmt(inst, 2), fmt(inst*(i+1), 2)] for i, m in enumerate(pm)]
        d.table(sched, [3.0, 3.6, 4.8, 4.9], 'Option B payment schedule', 0, font=9)
        d.p('Alternatives considered and not recommended are in section 13.4: a percentage of collections with floor and cap, and a hybrid of a fixed minimum plus a share of collections.')
        # ---------- 13 Returns ----------
        d.h1('13. Returns analysis and sensitivities')
        d.p('All returns are projections, labelled by case. IRRs are annual, compounded from quarterly flows; the investor funds in month ' + str(t0) + '.', italic=True)
        d.h2('13.1 Option A')
        ai = A['inv']
        d.table([['Case', 'Assumption', 'Investor IRR', 'Investor MOIC'], ['Base', 'Resale at final list price; 3% cost', pct(ai['Base']['irr']), mx(ai['Base']['moic'])], ['Downside', 'Resale 12 months later and 10% lower', pct(ai['Downside']['irr']), mx(ai['Downside']['moic'])], ['Upside', 'Resale 10% higher', pct(ai['Upside']['irr']), mx(ai['Upside']['moic'])],
                 ['Appreciation 40%', 'Launch to resale +40%', pct(ai['Appreciation 40%']['irr']), mx(ai['Appreciation 40%']['moic'])], ['Appreciation 60%', 'Launch to resale +60%', pct(ai['Appreciation 60%']['irr']), mx(ai['Appreciation 60%']['moic'])], ['No appreciation', 'Resale at launch list', pct(ai['Flat prices (0%)']['irr']), mx(ai['Flat prices (0%)']['moic'])]], [3.4, 6.9, 3.0, 3.0], 'Option A investor returns', 2, font=9.5)
        d.p(f'Break-even: the units can fall to {pct(D["breakeven"]["A_price_factor_moic1"],0)} of the final list price before the investor loses money on the ticket. The base case relies on list prices rising by {g_c*100:.0f}% (retail) and {g_a*100:.0f}% (offices); with no appreciation the units return {mx(ai["Flat prices (0%)"]["moic"])}.')
        d.table([['Face multiple', 'Investor IRR', 'Investor MOIC', 'Jalour cost of capital']] + [[mx(fx, 2), pct(v['irr']), mx(v['moic']), pct(v['jirr'])] for fx, v in A['sens_face'].items()], [3.4, 4.1, 4.1, 4.7], 'Sensitivity to the face multiple', 1, font=9.5)
        j = A['jal']
        d.p(f'**Cost to Jalour.** Giving the investor units from the first two sales tranches removes {pct(A["f"],0)} of those tranches\' collections. Jalour receives the ticket at funding and forgoes those collections (the units\' face value of EGP {fmt(A["face"],0)} million at list) as they would have been received, partly offset by commission saved. The landlord share is unchanged. On this basis Jalour\'s cost of capital is {pct(j["irr"])}; the NPV effect at 14% is EGP {fmt(j["npv"],1)} million and the project NPV after the deal is EGP {fmt(j["npv_project_after"],0)} million (before: {fmt(D["base"]["npv"],0)}). Peak cumulative shortage before the ticket rises to EGP {fmt(-j["peak_ex_ticket"],1)} million. If the same area were taken pro rata from all tranches (at higher average prices) the cost would be {pct(j["prorata_irr"])}.')
        d.figure(figp['inv'], 'Cumulative investor cash flow by option', 'Financial Annex, Option_A and Option_B sheets; projections')
        d.h2('13.2 Option B')
        d.table([['Measure', 'Base', 'Downside', 'Upside', 'Handover payments']] + [['Investor IRR if paid'] + [pct(S1[c]['irr']) for c in ('Base', 'Downside', 'Upside', 'Stress')], ['Investor MOIC'] + [mx(S1[c]['moic']) for c in ('Base', 'Downside', 'Upside', 'Stress')], ['Jalour cost of capital'] + [pct(S1[c]['jirr']) for c in ('Base', 'Downside', 'Upside', 'Stress')],
                 ['Lowest pooled cash after ticket and payouts (EGP m)'] + [fmt(S1[c]['min_cash'], 0) for c in ('Base', 'Downside', 'Upside', 'Stress')]], [6.0, 2.5, 2.6, 2.6, 2.6], 'Option B (fixed instalments)', 1, font=9.5)
        d.p(f'In the base case the lowest pooled cash of EGP {fmt(S1["Base"]["min_cash"],1)} million arises in month {S1["Base"]["min_cash_m"]}, before any instalment is due; it is the amount by which the modelled peak shortage exceeds the ticket and is funded by Jalour. Instalments are fixed, so the investor\'s IRR is the same in every case if Jalour pays. What changes is the cash available to pay: in the downside and handover cases pooled cash falls below zero and Jalour must fund the difference. Negative figures above are the amount Jalour would have to provide from sponsor equity or other capital sources.')
        d.h2('13.3 Payout coverage')
        net = D['series']['net']; pay = S1['Base']['pay']; liq = dict(S1['Base']['liq'])
        rows = [['Month', 'Net cash flow before payout (EGP m)', 'Instalment (EGP m)', 'Coverage: net cash flow / instalment', 'Coverage: cash available / instalment']]
        for m in pm:
            i = m//3-1; rows.append([str(m), fmt(net[i], 1), fmt(pay[i], 2), mx(net[i]/pay[i], 1), mx(liq[m], 1) if m in liq else 'n/a'])
        d.table(rows, [2.0, 4.0, 3.0, 3.7, 3.6], 'Payout coverage, base case', 1, font=9)
        d.p(f'Coverage on net cash flow alone is below the 1.5x threshold in {S1["Base"]["strict_below"]} of {S1["Base"]["n"]} quarters, because instalments fall in the construction peak and in the quarters after handover when collections are thin. "Cash available" adds pooled project cash, including the ticket and earlier collections, and is at least {mx(S1["Base"]["liq_min"],1)} in every quarter. Because nothing is ring-fenced, pooled cash is available to the project company but is also exposed to construction spending and to Jalour\'s own funding capacity. The investor should read the net-cash-flow measure as the stricter test.')
        d.h2('13.4 Alternatives considered')
        d.table([['Structure', 'Investor IRR (base)', 'Quarters < 1.5x (net cash flow)', 'Lowest pooled cash, downside (EGP m)', 'Last payment (month)'],
                 ['S1 Fixed multiple, equal instalments (recommended)', pct(S1['Base']['irr']), f'{S1["Base"]["strict_below"]} of {S1["Base"]["n"]}', fmt(S1['Downside']['min_cash'], 0), str(S1['Base']['last'])],
                 [f'S2 {pct(B["params"]["S2_pct"],1)} of collections, floor 0.4x, cap 1.5x of instalment', pct(S2['Base']['irr']), f'{S2["Base"]["strict_below"]} of {S2["Base"]["n"]}', fmt(S2['Downside']['min_cash'], 0), str(S2['Base']['last'])],
                 [f'S3 Fixed 1.8x plus {pct(B["params"]["S3_pct"],1)} of collections, cap 2.8x', pct(S3['Base']['irr']), f'{S3["Base"]["strict_below"]} of {S3["Base"]["n"]}', fmt(S3['Downside']['min_cash'], 0), str(S3['Base']['last'])]], [6.2, 2.5, 3.0, 2.8, 1.8], 'Option B structures compared', 1, font=9)
        d.p('S2 and S3 front-load payments into the large delivery-payment quarter and so show a higher IRR, but they pay by taking a share of collections that are themselves needed for construction, and they do not improve coverage in the construction peak. S1 is recommended for simplicity and certainty of the schedule. None of the three meets a 1.5x net-cash-flow test in every quarter; this is disclosed rather than cured by reducing the return, and it is the reason for Jalour\'s funding undertaking and the payment priority covenant.')
        d.h2('13.5 Project sensitivities')
        sp = D['sens']
        d.table([['Case', 'Net cash flow (EGP m)', 'NPV at 14% (EGP m)', 'Peak shortage (EGP m)'], ['Base', fmt(sp['price'][1.0]['net'], 0), fmt(sp['price'][1.0]['npv'], 0), fmt(-sp['price'][1.0]['peak'], 0)]] +
                [[f'Prices {int((f-1)*100):+d}%', fmt(sp['price'][f]['net'], 0), fmt(sp['price'][f]['npv'], 0), fmt(-sp['price'][f]['peak'], 0)] for f in (0.8, 0.9, 1.1, 1.2)] + [[f'Sales delayed {q*3} months', fmt(sp['delay'][q]['net'], 0), fmt(sp['delay'][q]['npv'], 0), fmt(-sp['delay'][q]['peak'], 0)] for q in (2, 4, 6, 8)] +
                [[f'Construction cost {int((f-1)*100):+d}%', fmt(sp['cost'][f]['net'], 0), fmt(sp['cost'][f]['npv'], 0), fmt(-sp['cost'][f]['peak'], 0)] for f in (1.1, 1.2)] + [[f'Collections slip {q} quarter{"s" if q > 1 else ""}', fmt(sp['slip'][q]['net'], 0), fmt(sp['slip'][q]['npv'], 0), fmt(-sp['slip'][q]['peak'], 0)] for q in (1, 2, 4)],
                [6.0, 3.4, 3.5, 3.4], 'Project sensitivities (Jalour net cash flow before any investor instrument)', 1, font=9)
        d.figure(figp['sens'], 'Project NPV by case', 'Financial Annex, Scenarios and Sensitivity sheets')
        # ---------- 14 Security ----------
        d.h1('14. Security and ranking')
        d.p('This is a proposal for review by Jalour\'s lawyers and the investor\'s counsel. Nothing in the structure is ring-fenced: there is no restricted account, escrow or segregated collection account, and security is contractual or over specific units.', keep=True)
        d.h2('14.1 Base package')
        d.table([['Protection', 'Option A', 'Option B'], ['Specific units', 'Registered allocation of identified units from the first two sales tranches; the units are the investor\'s asset', 'Earmark of unsold units with a list value of at least the ticket (substitution of equal value allowed); Jalour retained units not used'],
                 ['Payment priority covenant', 'Not applicable', 'No shareholder distributions or repayment of shareholder loans while an instalment is overdue'], ['Negative pledge', 'On the allocated units', 'On the earmarked units'], ['Information rights', 'Monthly report, annual audited accounts, audit right', 'Same'],
                 ['Sponsor funding undertaking', 'Funds project funding gaps; supports delivery', 'Same; not a guarantee of the return'], ['Landlord acknowledgement', 'Of the allocation and assignment rights', 'Of the earmark']], [3.6, 6.4, 6.3], 'Security and protections', -1, font=9)
        d.h2('14.2 Optional protections (for counsel)')
        d.bullets(['Promissory note or post-dated cheques for each Option B instalment, common in Egyptian practice.', 'Independent engineer reporting on construction progress and cost to complete.', 'Step-in right to market the earmarked units after a payment default.', 'Parent guarantee: not offered. Jalour has chosen not to guarantee the return; the investor should price this.'])
        d.h2('14.3 Ranking')
        d.p('Proposed ranking, for counsel: the investor\'s entitlement ranks as an obligation of the project company. It ranks behind amounts owed to the landowner under its agreement (which are paid from collections and take the first 35%) and behind construction and operating creditors in the ordinary course, and ahead of distributions to shareholders only to the extent of the payment priority covenant. Further project-level financing, if any, requires the investor\'s consent where it would rank ahead of the investor or encumber the earmarked units.')
        # ---------- 15 Governance ----------
        d.h1('15. Governance and information rights', page_break=False)
        d.bullets(['**Monthly report** within 15 days of month-end: units sold and value, collections against plan, costs against budget, construction progress against programme, cash position, any payment overdue to the landlord.', '**Quarterly review** with the investor, including cost to complete and the sales pipeline.', '**Annual audited accounts** of the project company within 120 days of year-end.', '**Audit and inspection rights** on reasonable notice over books, the sales register and the site.', '**Consent matters:** changes to the unit allocation or earmarked units, new project-level debt or security that would rank ahead, related-party transactions above an agreed threshold, and amendment of the landlord agreement that changes payments.', '**Notice events:** landlord dispute or default, loss or change of a permit, delay of more than three months to the construction programme, any payment overdue to the investor.'])
        d.p('Reporting thresholds and timings are indicative. Because there is no restricted account, the investor\'s monitoring rests on reporting and audit rather than on control of cash.')
        # ---------- 16 Exit ----------
        d.h1('16. Exit and liquidity', page_break=False)
        d.h2('16.1 Option A')
        d.bullets([f'Handover is in project month {I["delivery"]}; units can be registered and sold from handover. The base case assumes sale over four quarters from month {I["delivery"]+3}.', 'Before handover the investor may assign the contract with Jalour\'s consent and the landlord\'s acknowledgement [transfer fee to be agreed]. The buyer takes handover risk and will price it.', 'Jalour continues to sell its own inventory; the investor\'s resale competes with it. The agreements may include an orderly-sales covenant (for example, no more than an agreed share of the allocated area marketed in any quarter).'])
        d.table([['Assignment at month ' + str(t0+24), 'Investor IRR', 'Investor MOIC']] + [[f'Discount to face {int(dsc*100)}%', pct(v['irr']), mx(v['moic'])] for dsc, v in A['early'].items()], [7.0, 4.5, 4.8], 'Early assignment of the contract at month ' + str(t0+24) + ' (units valued at face, 3% cost)', 1, font=9.5)
        d.h2('16.2 Option B')
        d.p(f'Cash is returned through the instalments in section 12.2. There is no market for the entitlement. Transfer of the entitlement to a third party requires Jalour\'s consent. Payback (cumulative cash returned equals the ticket) occurs in month {D["B_payback"]}; for Option A, payback in the base case is month {D["A_payback"]}.')
        # ---------- 17 Risks ----------
        d.h1('17. Risk factors')
        d.p('The following is not exhaustive. Each risk is stated with the main mitigant.', keep=True)
        risks = [('Market and price risk', f'Unit prices may rise less than the plan, or fall. The plan needs list prices to rise by {g_c*100:.0f}% (retail) and {g_a*100:.0f}% (offices). Option A returns depend on resale value; with no appreciation they fall to {mx(ai["Flat prices (0%)"]["moic"])}. Mitigant: downside case shown; Option B has no unit price exposure.'),
                 ('Sales pace and absorption', 'The pipeline of Cairo offices equals 82% of current stock by 2029 (Knight Frank). If sales are delayed 12 months the project NPV falls to EGP ' + fmt(sp['delay'][4]['npv'], 0) + ' million and the peak funding need rises to EGP ' + fmt(-sp['delay'][4]['peak'], 0) + ' million. Mitigant: Jalour funding undertaking; phased sales.'),
                 ('Inflation and cost risk', 'Construction cost is flat in the model. A 10% overrun lowers NPV by EGP ' + fmt(D['base']['npv']-sp['cost'][1.1]['npv'], 0) + ' million. Mitigant: 5% contingency; contractor pricing in the data room; fixed-price contracting to be sought.'),
                 ('Currency risk', 'Returns are in EGP. USD value falls if the pound weakens; foreign investors should assess FX and repatriation (section 18). The USD equivalents shown use EGP 48 per USD; the market rate in September 2026 was about EGP 50 to 51.'),
                 ('Delivery and construction risk', 'The development must be built in 36 months from month ' + str(I['con_start']) + '. Delay delays resale of units and the collection of delivery payments. Mitigant: independent engineer (optional); reporting; Jalour funding undertaking.'),
                 ('Delivery-payment timing', 'The plan collects delivery payments ahead of handover under the sale contracts. If buyers do not pay until handover, the peak funding need is EGP ' + fmt(F['stress_peak'], 0) + ' million. Mitigant: contractual milestone; Jalour funding.'),
                 ('Landlord risk', 'The landlord\'s guarantee is payable regardless of sales and its consent is required for allocations and encumbrances. A dispute could delay or prevent the allocation. Mitigant: consent as a condition precedent; guarantee schedule funded in the plan.'),
                 ('Legal and regulatory risk', 'Rules on off-plan sales on state-linked land (including Decree 2184 of 2022), on assignment of contracts, on private placements and on foreign ownership may affect structure, timing or cost. Mitigant: counsel review before signing.'),
                 ('Counterparty and credit risk (Option B)', 'Instalments are an obligation of the project company without a parent guarantee. In the downside case pooled cash is negative and payment depends on Jalour\'s support. Mitigant: payment priority covenant; earmarked units; funding undertaking; optional promissory notes.'),
                 ('Concentration', 'The investment is exposed to a single project, in a single city, in two property types. There is no cross-collateral or cross-default with other projects.'),
                 ('Model and information risk', 'Figures derive from a Jalour model that has known simplifications: the admin sales area exceeds office built-up area, prices are stepped by assumption, costs are not inflated and pre-construction costs are not separately modelled. Mitigant: reconciliation and data room evidence.'),
                 ('Liquidity risk', 'There is no market for the Option B entitlement and Option A units cannot be sold in practice before handover without a discount.')]
        for nm, tx in risks:
            d.h3(nm); d.p(tx)
        # ---------- 18 Tax and legal ----------
        d.h1('18. Tax and legal considerations', page_break=False)
        d.p('The following are issues for counsel and tax advisers. Jalour gives no legal or tax conclusion.', keep=True)
        d.bullets(['Nature of the instrument (equity, loan or contractual entitlement) and its tax, regulatory and enforceability consequences in Egypt.', 'Regulatory position of a private placement of this kind (Financial Regulatory Authority and any other body), including investor eligibility.', 'Tax on the investor\'s return under each option: withholding tax, corporate or personal income tax, and VAT on supplies.', 'Taxes and registration fees on the transfer of units (including real estate disposal tax and stamp duty) and who bears them on resale.', 'Whether a separate project bank account is required for these plots under the rules for developers on state-linked land.', 'Foreign investor treatment: ownership of units, repatriation of proceeds, foreign-exchange availability and any central bank requirements.', 'Enforceability of security over unsold units, the earmark, negative pledge and any promissory notes; stamp duty on security documents.', 'Anti-money-laundering and know-your-customer requirements for the investor and for unit buyers.', 'Corporate approvals, project company formation and the effect of Jalour or Al Jalal Holding insolvency on the project company.'])
        # ---------- Appendices ----------
        d.h1('Appendices')
        d.h2('Appendix A: Definitions')
        d.table([['Term', 'Meaning'], ['Launch list', f'Price list in force at the first sales tranche: retail EGP {fmt(pl_c*1000,0)}, offices EGP {fmt(pl_a*1000,0)} per m2'], ['Face value', 'Units valued at the launch list'], ['Pooled cash', 'Cumulative project net cash flow plus the ticket less payouts; not a restricted account'], ['Coverage (net cash flow)', 'Quarterly project net cash flow before payouts divided by the instalment'], ['Coverage (cash available)', 'Pooled cash at the start of the quarter plus the quarter\'s net cash flow, divided by the instalment'], ['Peak shortage', 'Lowest cumulative net cash flow in the model, before the ticket'], ['MOIC', 'Multiple of invested capital: cash received divided by cash invested']], [4.5, 11.8], 'Definitions', -1, bold_first_col=True, font=9.5)
        d.h2('Appendix B: Reconciliation to the Financial Annex')
        d.p('Each headline figure in this memorandum is computed in the Financial Annex (Summary sheet). The Reconciliation sheet of the annex ties the project cash flow to the Jalour model (net cash flow, NPV, peak shortage, collections, landlord payments and costs) and to an independent recomputation; all differences are zero.')
        d.table([['Figure', 'Value', 'Annex location'], ['Net cash flow, nominal (EGP m)', fmt(D['base']['net'], 1), 'Summary; Project_CF'], ['NPV at 14% (EGP m)', fmt(D['base']['npv'], 1), 'Summary; Project_CF'], ['Peak shortage (EGP m)', fmt(F['model_peak'], 1), 'Summary; Project_CF'], ['Total uses (EGP m)', fmt(F['total'], 1), 'Sources_Uses'],
                 ['Option A investor IRR / MOIC', f'{pct(ai["Base"]["irr"])} / {mx(ai["Base"]["moic"])}', 'Option_A'], ['Option A Jalour cost of capital', pct(j['irr']), 'Option_A'], ['Option B investor IRR / MOIC', f'{pct(S1["Base"]["irr"])} / {mx(S1["Base"]["moic"])}', 'Option_B'], ['Option B lowest pooled cash, base (EGP m)', fmt(S1['Base']['min_cash'], 1), 'Option_B'], ['Downside peak shortage (EGP m)', fmt(F['downside_peak'], 0), 'Scenarios']], [7.5, 4.0, 4.8], 'Key figures and annex locations', 1, font=9.5)
        d.h2('Appendix C: Financial Annex contents')
        d.p('Cover, Summary, Inputs, Source_Data, Project_CF, Sources_Uses, Option_A, Option_B, Scenarios, Sensitivity, Reconciliation. Inputs are in blue; every calculation is a formula.')
        d.h2('Appendix D: Market sources')
        d.table([['Source', 'Used for']] + [[f['src'], f['text'][:110] + ('...' if len(f['text']) > 110 else '')] for f in MK['facts']], [6.0, 10.3], 'Market sources (reported; not independently verified)', -1, font=8.5)
        d.save(path); return d
    d1 = one(); d1.save(path)
    pdf, pages = pdf_pages(path); toc = find_pages(d1.headings, pages)
    toc_map = {t: toc.get(t.replace('Appendices', 'Appendices'), '') for t in []}
    # map TOC labels to heading text
    labels = {}
    for lvl, h in d1.headings:
        if lvl == 1: labels[h] = toc.get(h, '')
    d2 = one(labels); d2.save(path)
    pdf, pages = pdf_pages(path)
    return dict(pages=len(pages), pdf=pdf)
if __name__ == '__main__':
    r = build_memo(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4]); print(r)
