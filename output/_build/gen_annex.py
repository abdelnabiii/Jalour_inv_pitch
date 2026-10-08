"""Formula-driven Financial Annex generator. make_annex(k, T, path). Every calculation cell is a formula; inputs are blue."""
import sys; sys.path.insert(0, '.')
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.comments import Comment
import engine as E
from engine2 import *
import packdata as PD
NQ_ = 60
AR = 'Arial'
F_IN = Font(name=AR, size=10, color='0000FF'); F_N = Font(name=AR, size=10); F_B = Font(name=AR, size=10, bold=True); F_T = Font(name=AR, size=14, bold=True, color='1F3864')
F_H = Font(name=AR, size=10, bold=True, color='FFFFFF'); F_NOTE = Font(name=AR, size=9, italic=True, color='666666'); F_LINK = Font(name=AR, size=10, color='007A33')
FILL_H = PatternFill('solid', fgColor='1F3864'); FILL_K = PatternFill('solid', fgColor='E8EEF7'); FILL_OK = PatternFill('solid', fgColor='E2EFDA'); FILL_Y = PatternFill('solid', fgColor='FFF2CC')
N1 = '#,##0.0;(#,##0.0);-'; N2 = '#,##0.00;(#,##0.00);-'; PC = '0.0%'; X = '0.00"x"'
def qc(i): return CL(3+i)           # quarter i (1..60) -> column D..BK
FIRST, LAST = qc(1), qc(NQ_)
def rng(row, sheet=None): return (f"'{sheet}'!" if sheet else '') + f"${FIRST}${row}:${LAST}${row}"
def make_annex(k, T, path):
    I = PD.INFO[k]; p = P[k]; Lh = launch(k); b = series(k); name = I['name']
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    wb.properties.title = f'Financial Annex - {name} - Ticket {T}m'; wb.properties.creator = 'Jalour Developments'; wb.properties.subject = 'Confidential'; wb.properties.keywords = ''
    def sh(n, title, sub=None):
        ws = wb.create_sheet(n); ws['A1'] = title; ws['A1'].font = F_T
        if sub: ws['A2'] = sub; ws['A2'].font = F_NOTE
        ws.sheet_view.showGridLines = False; return ws
    def put(ws, ref, v, font=F_N, fmt=None, fill=None):
        c = ws[ref]; c.value = v; c.font = font
        if fmt: c.number_format = fmt
        if fill: c.fill = fill
        return c
    def hdr(ws, row, labels, col0=1):
        for j, h in enumerate(labels):
            c = ws.cell(row, col0+j, h); c.font = F_H; c.fill = FILL_H; c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    # ============ Inputs ============
    wsI = sh('Inputs', f'Inputs - {name}, ticket EGP {T}m (EGP million unless stated)', 'Blue = input. Series from the source model are on Source_Data. Every figure traces to a source cell or is labelled ASSUMPTION.')
    hdr(wsI, 4, ['Parameter', 'Value', 'Unit', 'Source / basis'])
    ref = {}; r = 5
    def inp(key, label, val, unit, src, fmt=None, formula=False):
        nonlocal r
        wsI.cell(r, 1, label).font = F_N
        c = wsI.cell(r, 2, val); c.font = F_N if formula else F_IN
        if fmt: c.number_format = fmt
        wsI.cell(r, 3, unit).font = F_N; wsI.cell(r, 4, src).font = F_NOTE
        ref[key] = f"Inputs!$B${r}"; r += 1
    def sec(t):
        nonlocal r
        r += 1; wsI.cell(r, 1, t).font = F_B; wsI.cell(r, 1).fill = FILL_K; r += 1
    sec('Transaction')
    inp('T', 'Investor ticket', T, 'EGP m', 'Offer term', N1)
    inp('face_x', 'Option A: face value multiple of ticket (units at launch list price)', FACE_X, 'x', 'Offer term (Option A)', X)
    inp('B_mult', 'Option B: total cash return multiple', M_B2, 'x', 'Offer term (Option B)', X)
    inp('grace', 'Option B: grace period', GRACE_M, 'months', 'Offer term', '0')
    inp('n_pay', 'Option B: number of quarterly payouts', PAY_Q, 'quarters', 'Offer term (3 years)', '0')
    inp('t0', 'Investor funding month (t0)', p['t0'], 'month', f"ASSUMPTION: month the down payment falls due in the model ('Net Cash Flow -With DP' row {11 if k=='GS' else 12})", '0')
    inp('handover', 'Project handover month', I['delivery'], 'month', f"Source model, {name} Summary!E4", '0')
    inp('disc', 'Discount rate', 0.14, '% p.a.', "Source model, sales sheet cell C72 (labelled Cost Of Capital); kept at 14% by Jalour", PC)
    inp('fric', 'Option A: investor resale friction', FRICTION, '%', 'ASSUMPTION: broker plus developer transfer fee', PC)
    inp('real_q', 'Option A: resale spread (quarters from handover + 3 months)', REAL_Q, 'quarters', 'ASSUMPTION', '0')
    sec('Project and landlord terms (source model)')
    inp('ll', 'Landlord share of collections', LL, '%', 'Offers DP!C5', PC)
    inp('minG', 'Landlord minimum guarantee', p['minG'], 'EGP m', 'Offers DP!C3' if k == 'GS' else 'Offers DP!D3', N1)
    inp('dp', 'Down payment to landlord (10% of guarantee)', p['dp'], 'EGP m', 'Offers DP!C9' if k == 'GS' else 'Offers DP!D9', N1)
    inp('list100', 'Total sales value at 100% of units', p['list100'], 'EGP m', 'Summary C27 (sold 80% plus retained 20% at modelled prices)', N1)
    inp('comm_rate', 'Sales commission rate', COMM, '%', 'Summary B19', PC)
    inp('prof_fee', 'Professional fees (% of construction cost)', 0.05, '%', "Cash Out Detail!C53", PC)
    inp('precon_share', 'Share of professional fees incurred before construction start', 0.30, '%', 'ASSUMPTION: the model has no pre-construction cost; to be confirmed', PC)
    sec('Launch price list and tranche values (source model sales sheet)')
    inp('pl_c', 'Launch list price, retail', Lh['p0'][0], 'EGP k / m2', 'First sales column, row 3', '#,##0')
    inp('pl_a', 'Launch list price, offices', Lh['p0'][1], 'EGP k / m2', 'First sales column, row 4', '#,##0')
    inp('pf_c', 'Delivery list price, retail (last list price in the plan)', Lh['fin'][0], 'EGP k / m2', 'Last sales column, row 3', '#,##0')
    inp('pf_a', 'Delivery list price, offices (last list price in the plan)', Lh['fin'][1], 'EGP k / m2', 'Last sales column, row 4', '#,##0')
    inp('Vc', 'Launch tranches: retail sales value at list', Lh['Vc'], 'EGP m', f"First two sales tranches (months {Lh['months'][0]} and {Lh['months'][1]}), row 24", N1)
    inp('Va', 'Launch tranches: office sales value at list', Lh['Va'], 'EGP m', 'Same tranches, row 25', N1)
    inp('appr', 'Price appreciation launch to delivery list (value-weighted)', f"=({ref['Vc'].split('!')[1]}/({ref['Vc'].split('!')[1]}+{ref['Va'].split('!')[1]}))*({ref['pf_c'].split('!')[1]}/{ref['pl_c'].split('!')[1]})+({ref['Va'].split('!')[1]}/({ref['Vc'].split('!')[1]}+{ref['Va'].split('!')[1]}))*({ref['pf_a'].split('!')[1]}/{ref['pl_a'].split('!')[1]})", 'x', 'Calculated. Investor resale value = face x this factor (base)', '0.0000', True)
    sec('Option B alternatives (considered, not recommended)')
    pc = calibrate_min_pct(k, T, M_B2, b, 1.5*M_B2*T/12, 0.4*M_B2*T/12); ph = calibrate_hybrid(k, T, 1.8, M_B2, 2.8, b)
    inp('s2_pct', 'S2: share of collections', pc, '%', 'Calibrated (Python) so the base case reaches the target multiple by the end of the payout window', '0.0000%')
    inp('s2_floor', 'S2: floor per quarter as multiple of equal instalment', 0.4, 'x', 'Design', '0.00')
    inp('s2_cap', 'S2: cap per quarter as multiple of equal instalment', 1.5, 'x', 'Design', '0.00')
    inp('s3_fixed', 'S3: fixed minimum multiple', 1.8, 'x', 'Design', X)
    inp('s3_cap', 'S3: total cap multiple', 2.8, 'x', 'Design', X)
    inp('s3_pct', 'S3: share of collections', ph, '%', 'Calibrated (Python)', '0.0000%')
    inp('gate', 'Coverage test threshold', 1.5, 'x', 'Brief', X)
    sec('Reference values from the source model (for reconciliation only)')
    m = PD.MODEL_REF[k]
    inp('ref_net', 'Jalour net cash flow, nominal', m['net'], 'EGP m', "'Net Cash Flow -With DP'!C46" if k == 'GS' else "'Net Cash Flow -With DP'!C47", '#,##0.000000')
    inp('ref_npv', 'NPV of net cash flow with down payment', m['npv'], 'EGP m', 'Summary!D10' if k == 'GS' else 'Summary!D11', '#,##0.000000')
    inp('ref_peak', 'Peak cumulative cash shortage', m['peak'], 'EGP m', f"{name} Summary!E7", '#,##0.000000')
    inp('ref_coll', 'Total collections (80% of units sold)', m['coll'], 'EGP m', "'Net Cash Flow -With DP'!C4" if k == 'GS' else "'Net Cash Flow -With DP'!C5", '#,##0.000000')
    inp('ref_land', 'Total landlord payments', m['land'], 'EGP m', "'Net Cash Flow -With DP'!C11" if k == 'GS' else "'Net Cash Flow -With DP'!C12", '#,##0.000000')
    inp('ref_cost', 'Total construction, commission and SG&A', m['cost'], 'EGP m', "'Net Cash Flow -With DP'!C32" if k == 'GS' else "'Net Cash Flow -With DP'!C33", '#,##0.000000')
    inp('usd', 'USD exchange rate (executive summary only)', 48.0, 'EGP per USD', 'Brief / model', '0.00')
    for col, w in zip('ABCD', (66, 16, 14, 100)): wsI.column_dimensions[col].width = w
    R = lambda key: ref[key]
    # derived: first payout month
    r += 1; wsI.cell(r, 1, 'First payout month').font = F_N
    wsI.cell(r, 2, f"={R('t0')}+{R('grace')}+3").font = F_N; wsI.cell(r, 3, 'month'); wsI.cell(r, 4, 'Calculated: funding month + grace + 3').font = F_NOTE; ref['first_pay'] = f"Inputs!$B${r}"; r += 1
    wsI.cell(r, 1, 'Launch tranche face capacity (value at launch list)').font = F_N
    wsI.cell(r, 2, f"={R('Vc')}+{R('Va')}").font = F_N; wsI.cell(r, 2).number_format = N1; wsI.cell(r, 3, 'EGP m'); wsI.cell(r, 4, 'Calculated: retail plus office value of the launch tranches').font = F_NOTE; ref['Vl'] = f"Inputs!$B${r}"; r += 1
    # ============ Source_Data ============
    wsS = sh('Source_Data', f'Source model series - {name} (EGP million, quarter-end months)', 'Blue = values transcribed from the source model cells named in column B. Columns D onward = quarters 1..60 (month = 3 x quarter).')
    hdr(wsS, 4, ['Series', 'Source cell / basis', 'Total'] + [''] * NQ_)
    S = {}
    srows = [('idx', 'Quarter index', 'structural'), ('month', 'Month', '3 x index'), ('coll_ex', 'Collections excluding delivery payments', "Net Cash Flow -With DP row %d less delivery component" % (4 if k == 'GS' else 5)),
             ('deliv', 'Delivery payments (within collections)', f"{name} sheet rows 35-36"), ('sched', 'Landlord payments: down payment and guarantee schedule', f"Net Cash Flow -With DP row {11 if k=='GS' else 12}, schedule part"),
             ('exc', 'Landlord payments: excess over guarantee (years 8+)', 'Same row, excess settlement'), ('cons', 'Construction cost', 'Cash Out Detail row %d' % (28 if k == 'GS' else 29)),
             ('comm', 'Sales commission', 'Cash Out Detail row %d' % (34 if k == 'GS' else 35)), ('sga', 'SG&A and others (15% of construction)', 'Cash Out Detail row %d' % (40 if k == 'GS' else 41)),
             ('lf_all', 'Launch tranches: collections, retail and offices', f"{name} sheet rows 31-36, 38-69 for the first two sales tranches"), ('lf_c', 'Launch tranches: collections, retail', 'same, commercial rows'), ('lf_a', 'Launch tranches: collections, offices', 'same, admin rows'),
             ('lsv', 'Launch tranches: sales value at the quarter commission is paid (12 months after sale)', 'Cash Out Detail row 34 timing')]
    coll_ex = (b['coll'] - pad(E.D[k]['delivery_pay'])); deliv = pad(E.D[k]['delivery_pay'])
    lsv = np.zeros(NQ)
    for t in Lh['tr']:
        lsv[(t['month']+12)//3-1] += (t['sales_comm']+t['sales_admin'])/1000
    vals = {'coll_ex': coll_ex, 'deliv': deliv, 'sched': b['sched'], 'exc': b['exc'], 'cons': b['cons'], 'comm': b['comm'], 'sga': b['sga'], 'lf_all': Lh['flow']['all'], 'lf_c': Lh['flow']['comm'], 'lf_a': Lh['flow']['admin'], 'lsv': lsv}
    for n_, (key, label, srcn) in enumerate(srows):
        row = 5+n_; S[key] = row; wsS.cell(row, 1, label).font = F_N; wsS.cell(row, 2, srcn).font = F_NOTE
        for i in range(1, NQ_+1):
            c = wsS[f'{qc(i)}{row}']
            if key == 'idx': c.value = i; c.font = F_N
            elif key == 'month': c.value = f'={qc(i)}{S["idx"]}*3'; c.font = F_N
            else: c.value = float(vals[key][i-1]); c.font = F_IN; c.number_format = N2
        if key not in ('idx', 'month'):
            wsS.cell(row, 3, f'=SUM({FIRST}{row}:{LAST}{row})').number_format = N2; wsS.cell(row, 3).font = F_B
    wsS.column_dimensions['A'].width = 62; wsS.column_dimensions['B'].width = 46; wsS.column_dimensions['C'].width = 12
    for i in range(1, NQ_+1): wsS.column_dimensions[qc(i)].width = 9
    wsS.freeze_panes = 'D5'
    SR = lambda key: rng(S[key], 'Source_Data')
    def SC(key, i): return f"Source_Data!{qc(i)}${S[key]}"
    # ============ Project_CF (base) ============
    wsP = sh('Project_CF', f'{name}: project cash flow to Jalour, base case (quarterly, as per the source model)', 'Net cash flow = collections less landlord payments, construction, commission and SG&A. Positive = cash in.')
    hdr(wsP, 4, ['Line (EGP m)', 'Unit', 'Total'] + [''] * NQ_)
    P_ = {n: 5+j for j, n in enumerate(['idx', 'month', 't', 'coll', 'land', 'cons', 'comm', 'sga', 'net', 'cum', 'df', 'pv'])}
    labels = {'idx': 'Quarter', 'month': 'Month', 't': 'Years from month 0', 'coll': 'Collections (80% of units sold)', 'land': 'Landlord payments (35% share, guarantee, down payment)', 'cons': 'Construction cost', 'comm': 'Sales commission', 'sga': 'SG&A and others',
              'net': 'Net cash flow', 'cum': 'Cumulative net cash flow', 'df': 'Discount factor', 'pv': 'Present value of net cash flow'}
    for key, row in P_.items():
        wsP.cell(row, 1, labels[key]).font = F_B if key in ('net', 'cum') else F_N
        for i in range(1, NQ_+1):
            col = qc(i); c = wsP[f'{col}{row}']
            c.value = {'idx': f'={SC("idx", i)}', 'month': f'={SC("month", i)}', 't': f'={col}{P_["month"]}/12', 'coll': f'={SC("coll_ex", i)}+{SC("deliv", i)}', 'land': f'={SC("sched", i)}+{SC("exc", i)}',
                       'cons': f'={SC("cons", i)}', 'comm': f'={SC("comm", i)}', 'sga': f'={SC("sga", i)}', 'net': f'={col}{P_["coll"]}-{col}{P_["land"]}-{col}{P_["cons"]}-{col}{P_["comm"]}-{col}{P_["sga"]}',
                       'cum': f'={col}{P_["net"]}' if i == 1 else f'={qc(i-1)}{P_["cum"]}+{col}{P_["net"]}', 'df': f'=1/(1+{R("disc")})^{col}{P_["t"]}', 'pv': f'={col}{P_["net"]}*{col}{P_["df"]}'}[key]
            c.font = F_LINK if key in ('coll', 'land', 'cons', 'comm', 'sga') else F_N
            c.number_format = '0.0000' if key == 'df' else ('0.00' if key == 't' else ('0' if key in ('idx', 'month') else N2))
        if key in ('coll', 'land', 'cons', 'comm', 'sga', 'net', 'pv'):
            wsP.cell(row, 3, f'=SUM({FIRST}{row}:{LAST}{row})').number_format = N2; wsP.cell(row, 3).font = F_B
    pr = P_['net']+5
    for j, (lab, f_, fm) in enumerate([('Net cash flow, nominal', f'=C{P_["net"]}', N2), ('NPV of net cash flow @ discount rate', f'=C{P_["pv"]}', N2), ('Peak cumulative cash shortage', f'=MIN({FIRST}{P_["cum"]}:{LAST}{P_["cum"]})', N2),
                                       ('Month of peak shortage', f'=INDEX({FIRST}{P_["month"]}:{LAST}{P_["month"]},MATCH(B{pr+2},{FIRST}{P_["cum"]}:{LAST}{P_["cum"]},0))', '0')]):
        wsP.cell(pr+j, 1, lab).font = F_B; c = wsP.cell(pr+j, 2, f_); c.number_format = fm; c.font = F_B; c.fill = FILL_K
    PK = {'net': f'Project_CF!$B${pr}', 'npv': f'Project_CF!$B${pr+1}', 'peak': f'Project_CF!$B${pr+2}', 'peak_m': f'Project_CF!$B${pr+3}'}
    wsP.column_dimensions['A'].width = 56; wsP.column_dimensions['B'].width = 14; wsP.column_dimensions['C'].width = 12
    for i in range(1, NQ_+1): wsP.column_dimensions[qc(i)].width = 9
    wsP.freeze_panes = 'D5'
    PR = lambda key: rng(P_[key], 'Project_CF')
    def PC_(key, i): return f"Project_CF!{qc(i)}${P_[key]}"
    # ============ Sources_Uses ============
    wsU = sh('Sources_Uses', f'{name}: sources and uses of funds at the modelled peak, ticket EGP {T}m', 'Nothing is ring-fenced: the ticket is project funding, not a restricted account.')
    hdr(wsU, 4, ['Uses (EGP m)', 'Amount', 'Basis'])
    U = {}
    rows = [('dp', 'Down payment to landlord', f"={R('dp')}", 'Source model'), ('guar', 'Landlord guarantee instalments and working capital to modelled peak, net of collections', f"=-{PK['peak']}-{R('dp')}", 'Peak shortage less down payment'),
            ('pre', 'Design, permits and pre-construction (ASSUMPTION)', f"={R('prof_fee')}*Project_CF!$C${P_['cons']}*{R('precon_share')}", '30% of the 5% professional-fee line, incurred ahead of construction start'),
            ('tot', 'Total uses at peak', '=SUM(B5:B7)', '')]
    for j, (key, lab, f_, bs) in enumerate(rows):
        rr = 5+j; U[key] = rr; wsU.cell(rr, 1, lab).font = F_B if key == 'tot' else F_N; c = wsU.cell(rr, 2, f_); c.number_format = N1; c.font = F_B if key == 'tot' else F_N; wsU.cell(rr, 3, bs).font = F_NOTE
    hdr(wsU, 11, ['Sources (EGP m)', 'Amount', 'Basis'])
    src = [('tk', 'Investor ticket', f"={R('T')}", 'This offer'), ('bal', 'Balance of project funding', f"=MAX(0,B8-B12)", 'The balance of project funding is provided by Jalour sponsor equity, project collections and other capital sources.'),
           ('src', 'Total sources', '=B12+B13', ''), ('buf', 'Ticket in excess of modelled uses (general project liquidity, not ring-fenced)', '=MAX(0,B12-B8)', 'Flag if above zero'),
           ('sh1', 'Ticket as % of modelled peak shortage', f"=B12/-{PK['peak']}", ''), ('sh2', 'Ticket applied to uses as % of total uses', '=MIN(B12,B8)/B8', ''), ('xs', 'Ticket above modelled peak shortage', f"=MAX(0,B12+{PK['peak']})", 'Flag if above zero')]
    for j, (key, lab, f_, bs) in enumerate(src):
        rr = 12+j; U[key] = rr; wsU.cell(rr, 1, lab).font = F_N; c = wsU.cell(rr, 2, f_); c.font = F_N; c.number_format = PC if key in ('sh1', 'sh2') else N1; wsU.cell(rr, 3, bs).font = F_NOTE
    wsU.cell(20, 1, 'Funding-gap exposure (sponsor covers gaps in excess of project cash)').font = F_B
    wsU.column_dimensions['A'].width = 80; wsU.column_dimensions['B'].width = 16; wsU.column_dimensions['C'].width = 110
    # ============ scenario block writer ============
    def block(ws, r0, title, price=1.0, delay=0, cost=1.0, slip=0, deliv=0):
        ws.cell(r0, 1, title).font = F_B; ws.cell(r0, 1).fill = FILL_K
        names = ['Price factor', 'Sales delay (quarters)', 'Cost factor', 'Collections slip (quarters)', 'Delivery payments at handover (1 = yes)']
        pcells = []
        for j, (n_, v_) in enumerate(zip(names, (price, delay, cost, slip, deliv))):
            ws.cell(r0+1+j, 1, n_).font = F_N; c = ws.cell(r0+1+j, 2, v_); c.font = F_IN; pcells.append(f'$B${r0+1+j}')
        pr_, dl_, cs_, sl_, dv_ = pcells
        rows = ['raw', 'coll', 'land', 'cons', 'comm', 'sga', 'net', 'cum', 'a', 'bpay', 'binv', 'cash', 'cov1', 'cov2']
        lab = {'raw': 'Collections before delay and price', 'coll': 'Collections', 'land': 'Landlord payments', 'cons': 'Construction cost', 'comm': 'Sales commission', 'sga': 'SG&A and others', 'net': 'Net cash flow',
               'cum': 'Cumulative net cash flow', 'a': 'Option A investor cash flow', 'bpay': 'Option B payouts (fixed schedule)', 'binv': 'Option B investor cash flow', 'cash': 'Pooled project cash after ticket and payouts',
               'cov1': 'Coverage: net cash flow / payout', 'cov2': 'Coverage: cash available / payout'}
        BR = {n_: r0+7+j for j, n_ in enumerate(rows)}
        for n_, rr in BR.items(): ws.cell(rr, 1, lab[n_]).font = F_B if n_ in ('net', 'cum') else F_N
        ws.cell(r0+6, 1, 'Resale start month (Option A)').font = F_N
        ws.cell(r0+6, 2, f"={R('handover')}+3+3*{dl_}").font = F_N; start = f'$B${r0+6}'
        exc_tot = f"Source_Data!$C${S['exc']}"
        for i in range(1, NQ_+1):
            col = qc(i); mo = f'Source_Data!{col}${S["month"]}'; ix = f'Source_Data!{col}${S["idx"]}'
            f = {}
            f['raw'] = f'={SC("coll_ex", i)}+IF({dv_}=1,IF({mo}={R("handover")},Source_Data!$C${S["deliv"]},0),{SC("deliv", i)})'
            f['coll'] = f'={pr_}*IF({ix}-({dl_}+{sl_})>=1,INDEX(${FIRST}${BR["raw"]}:${LAST}${BR["raw"]},1,{ix}-({dl_}+{sl_})),0)'
            f['land'] = f'={SC("sched", i)}+{SC("exc", i)}*IF({exc_tot}=0,0,MAX(0,{R("ll")}*{R("list100")}*{pr_}-{R("minG")})/{exc_tot})'
            f['cons'] = f'={SC("cons", i)}*{cs_}'; f['sga'] = f'={SC("sga", i)}*{cs_}'
            f['comm'] = f'={pr_}*IF({ix}-{dl_}>=1,INDEX({SR("comm")},1,{ix}-{dl_}),0)'
            f['net'] = f'={col}{BR["coll"]}-{col}{BR["land"]}-{col}{BR["cons"]}-{col}{BR["comm"]}-{col}{BR["sga"]}'
            f['cum'] = f'={col}{BR["net"]}' if i == 1 else f'={qc(i-1)}{BR["cum"]}+{col}{BR["net"]}'
            f['a'] = f'=IF({mo}={R("t0")},-{R("T")},0)+IF(AND({mo}>={start},{mo}<{start}+3*{R("real_q")},MOD({mo}-{start},3)=0),{R("face_x")}*{R("T")}*{R("appr")}*{pr_}*(1-{R("fric")})/{R("real_q")},0)'
            f['bpay'] = f'=IF(AND({mo}>={R("first_pay")},{mo}<{R("first_pay")}+3*{R("n_pay")}),{R("B_mult")}*{R("T")}/{R("n_pay")},0)'
            f['binv'] = f'=IF({mo}={R("t0")},-{R("T")},0)+{col}{BR["bpay"]}'
            f['cash'] = (f'=IF({mo}={R("t0")},{R("T")},0)+{col}{BR["net"]}-{col}{BR["bpay"]}' if i == 1 else f'={qc(i-1)}{BR["cash"]}+IF({mo}={R("t0")},{R("T")},0)+{col}{BR["net"]}-{col}{BR["bpay"]}')
            f['cov1'] = f'=IF({col}{BR["bpay"]}>0,{col}{BR["net"]}/{col}{BR["bpay"]},"")'
            prev = '0' if i == 1 else f'{qc(i-1)}{BR["cash"]}'
            f['cov2'] = f'=IF({col}{BR["bpay"]}>0,({prev}+{col}{BR["net"]}+IF({mo}={R("t0")},{R("T")},0))/{col}{BR["bpay"]},"")'
            for n_ in rows:
                c = ws[f'{col}{BR[n_]}']; c.value = f[n_]; c.font = F_N; c.number_format = '0.00' if n_ in ('cov1', 'cov2') else N2
        # outputs
        o = r0+7+len(rows)+1; O = {}
        outs = [('net', 'Net cash flow, nominal', f'=SUM({FIRST}{BR["net"]}:{LAST}{BR["net"]})', N1), ('npv', 'NPV @ discount rate', f'=SUMPRODUCT({FIRST}{BR["net"]}:{LAST}{BR["net"]},{PR("df")})', N1),
                ('peak', 'Peak cumulative cash shortage', f'=MIN({FIRST}{BR["cum"]}:{LAST}{BR["cum"]})', N1), ('peak_m', 'Month of peak', f'=INDEX({PR("month")},MATCH(MIN({FIRST}{BR["cum"]}:{LAST}{BR["cum"]}),{FIRST}{BR["cum"]}:{LAST}{BR["cum"]},0))', '0'),
                ('a_irr', 'Option A investor IRR', f'=(1+IRR({FIRST}{BR["a"]}:{LAST}{BR["a"]},0.05))^4-1', PC), ('a_moic', 'Option A investor MOIC', f'=SUMIF({FIRST}{BR["a"]}:{LAST}{BR["a"]},">0")/-SUMIF({FIRST}{BR["a"]}:{LAST}{BR["a"]},"<0")', X),
                ('b_irr', 'Option B investor IRR', f'=(1+IRR({FIRST}{BR["binv"]}:{LAST}{BR["binv"]},0.05))^4-1', PC), ('b_cash', 'Option B lowest pooled cash (negative = sponsor support needed)', f'=MIN({FIRST}{BR["cash"]}:{LAST}{BR["cash"]})', N1),
                ('b_cov1', 'Option B lowest quarterly coverage (net cash flow / payout)', f'=MIN({FIRST}{BR["cov1"]}:{LAST}{BR["cov1"]})', '0.00'), ('b_cov1n', 'Quarters below threshold (strict)', f'=COUNTIF({FIRST}{BR["cov1"]}:{LAST}{BR["cov1"]},"<"&{R("gate")})', '0'),
                ('b_cov2', 'Option B lowest cash-available coverage', f'=MIN({FIRST}{BR["cov2"]}:{LAST}{BR["cov2"]})', '0.00'), ('b_cov2n', 'Quarters below threshold (cash available)', f'=COUNTIF({FIRST}{BR["cov2"]}:{LAST}{BR["cov2"]},"<"&{R("gate")})', '0')]
        for j, (key, lab_, f_, fm) in enumerate(outs):
            ws.cell(o+j, 1, lab_).font = F_B; c = ws.cell(o+j, 2, f_); c.font = F_B; c.number_format = fm; c.fill = FILL_K; O[key] = f"'{ws.title}'!$B${o+j}"
        return o+len(outs)+2, O
    # ============ Scenarios & Sensitivity ============
    wsC = sh('Scenarios', f'{name}: scenario analysis (live blocks; edit the blue parameters)', 'Each block recalculates the project, Option A (investor resale at the scenario price and delay) and Option B (fixed payouts) from Source_Data.')
    wsC.column_dimensions['A'].width = 64; wsC.column_dimensions['B'].width = 14
    for i in range(1, NQ_+1): wsC.column_dimensions[qc(i)].width = 9
    SCN = {}; row = 4
    for nm, kw in PD.scen_defs().items():
        row, O = block(wsC, row, nm, kw.get('price', 1.0), kw.get('delay_q', 0), kw.get('cost', 1.0), kw.get('slip_q', 0), 1 if kw.get('deliv_aligned') else 0); SCN[nm] = O
    wsC.freeze_panes = 'C4'
    wsX = sh('Sensitivity', f'{name}: one-factor sensitivities (live blocks)', 'Sales delay, price change, cost overrun and collections slippage.')
    wsX.column_dimensions['A'].width = 64; wsX.column_dimensions['B'].width = 14
    for i in range(1, NQ_+1): wsX.column_dimensions[qc(i)].width = 9
    SEN = {}; row = 4
    for fam, lst in (('price', [(f, dict(price=f)) for f in (0.8, 0.9, 1.1, 1.2)]), ('delay', [(q, dict(delay_q=q)) for q in (2, 4, 6, 8)]), ('cost', [(f, dict(cost=f)) for f in (1.1, 1.2)]), ('slip', [(q, dict(slip_q=q)) for q in (1, 2, 4)])):
        for v, kw in lst:
            row, O = block(wsX, row, f'{fam}: {v}', kw.get('price', 1.0), kw.get('delay_q', 0), kw.get('cost', 1.0), kw.get('slip_q', 0), 0); SEN[(fam, v)] = O
    # sensitivity summary table at top is in Summary sheet
    # ============ Option_A ============
    wsA = sh('Option_A', f'{name}: Option A (units), ticket EGP {T}m', 'Units priced at launch list; allocated from the first two sales tranches; Jalour bears the landlord 35% on those units (landlord payments unchanged).')
    wsA.column_dimensions['A'].width = 66; wsA.column_dimensions['B'].width = 16
    for i in range(1, NQ_+1): wsA.column_dimensions[qc(i)].width = 9
    hdr(wsA, 4, ['Allocation', 'Value', 'Unit'])
    al = [('face', 'Face value of units at launch list', f"={R('face_x')}*{R('T')}", N1, 'EGP m'), ('share', 'Share of launch-tranche inventory allocated', f"=B5/{R('Vl')}", PC, ''),
          ('face_c', 'of which retail', f"=B5*{R('Vc')}/{R('Vl')}", N1, 'EGP m'), ('face_a', 'of which offices', f"=B5*{R('Va')}/{R('Vl')}", N1, 'EGP m'),
          ('area_c', 'Retail area allocated', f"=B7*1000/{R('pl_c')}", '#,##0', 'm2'), ('area_a', 'Office area allocated', f"=B8*1000/{R('pl_a')}", '#,##0', 'm2'),
          ('val', 'Value at delivery list price (before resale friction)', f"=B5*{R('appr')}", N1, 'EGP m'), ('ll', 'Landlord 35% on the allocated units, borne by Jalour', f"={R('ll')}*B5", N1, 'EGP m')]
    AR_ = {}
    for j, (key, lab, f_, fm, un) in enumerate(al):
        rr = 5+j; AR_[key] = rr; wsA.cell(rr, 1, lab).font = F_N; c = wsA.cell(rr, 2, f_); c.number_format = fm; c.font = F_N; wsA.cell(rr, 3, un)
    hdr(wsA, 15, ['Quarterly flows (EGP m)', 'Total'] + [''] * (NQ_+1))
    ar = {'month': 16, 'inv': 17, 'jal_t': 18, 'jal_c': 19, 'jal_s': 20, 'jal': 21, 'jcum': 22}
    wsA.cell(15, 3, '').font = F_H
    labs = {'month': 'Month', 'inv': 'Investor cash flow (base)', 'jal_t': 'Jalour: ticket received', 'jal_c': 'Jalour: collections forgone on allocated units', 'jal_s': 'Jalour: commission saved', 'jal': 'Jalour: net incremental cash flow', 'jcum': 'Jalour: pooled position effect (cumulative)'}
    for key, rr in ar.items(): wsA.cell(rr, 1, labs[key]).font = F_N
    for i in range(1, NQ_+1):
        col = qc(i); mo = f'Source_Data!{col}${S["month"]}'
        vals_ = {'month': f'={mo}', 'inv': f"=Scenarios!{col}{0}"}
        wsA[f'{col}{ar["month"]}'] = f'={mo}'
        wsA[f'{col}{ar["inv"]}'] = f"=IF({mo}={R('t0')},-{R('T')},0)+IF(AND({mo}>={R('handover')}+3,{mo}<{R('handover')}+3+3*{R('real_q')},MOD({mo}-({R('handover')}+3),3)=0),$B$5*{R('appr')}*(1-{R('fric')})/{R('real_q')},0)"
        wsA[f'{col}{ar["jal_t"]}'] = f"=IF({mo}={R('t0')},{R('T')},0)"
        wsA[f'{col}{ar["jal_c"]}'] = f"=-$B$6*{SC('lf_all', i)}"
        wsA[f'{col}{ar["jal_s"]}'] = f"=$B$6*{SC('lsv', i)}*{R('comm_rate')}"
        wsA[f'{col}{ar["jal"]}'] = f'={col}{ar["jal_t"]}+{col}{ar["jal_c"]}+{col}{ar["jal_s"]}'
        wsA[f'{col}{ar["jcum"]}'] = f'={col}{ar["jal"]}' if i == 1 else f'={qc(i-1)}{ar["jcum"]}+{col}{ar["jal"]}'
        for key, rr in ar.items(): wsA[f'{col}{rr}'].number_format = '0' if key == 'month' else N2; wsA[f'{col}{rr}'].font = F_N
    for key in ('inv', 'jal_t', 'jal_c', 'jal_s', 'jal'): wsA.cell(ar[key], 2, f'=SUM({FIRST}{ar[key]}:{LAST}{ar[key]})').number_format = N2
    hdr(wsA, 25, ['Returns', 'Value', 'Note'])
    inv_r = f'{FIRST}{ar["inv"]}:{LAST}{ar["inv"]}'; jal_r = f'{FIRST}{ar["jal"]}:{LAST}{ar["jal"]}'
    ret = [('irr', 'Investor IRR, base', f'=(1+IRR({inv_r},0.05))^4-1', PC, 'Appreciation to delivery list price in the model'), ('moic', 'Investor MOIC, base', f'=SUMIF({inv_r},">0")/-SUMIF({inv_r},"<0")', X, ''),
           ('payback', 'Investor payback month', f'=INDEX({FIRST}{ar["month"]}:{LAST}{ar["month"]},MATCH(1,{FIRST}27:{LAST}27,0))', '0', 'First quarter cumulative cash is not negative'),
           ('jirr', 'Jalour cost of capital (IRR of ticket vs forgone collections)', f'=(1+IRR({jal_r},0.05))^4-1', PC, 'Jalour bears landlord share'), ('jnpv', 'NPV to Jalour @ discount rate', f'=SUMPRODUCT({jal_r},{PR("df")})', N1, ''),
           ('jnom', 'Nominal net cost to Jalour', f'=B{ar["jal"]}', N1, ''), ('proj_npv', 'Project NPV after the deal', f"={PK['npv']}+B30", N1, 'Project NPV before the deal plus deal NPV'),
           ('peak', 'Peak cumulative shortage excluding ticket cash (after forgone collections)', f'=MIN({FIRST}38:{LAST}38)', N1, 'Project cumulative plus deal effect, ticket cash excluded')]
    for j, (key, lab, f_, fm, nt) in enumerate(ret):
        rr = 26+j; wsA.cell(rr, 1, lab).font = F_N; c = wsA.cell(rr, 2, f_); c.number_format = fm; c.font = F_B; c.fill = FILL_K; wsA.cell(rr, 3, nt).font = F_NOTE
    # payback helper row 27? place helper in row 35 and fix formula
    wsA.cell(35, 1, 'Investor cumulative cash flow').font = F_N; wsA.cell(36, 1, 'Cumulative not negative (1/0)').font = F_N
    for i in range(1, NQ_+1):
        col = qc(i); wsA[f'{col}35'] = f'={col}{ar["inv"]}' if i == 1 else f'={qc(i-1)}35+{col}{ar["inv"]}'; wsA[f'{col}36'] = f'=IF(AND({col}35>=0,{col}{ar["month"]}>{R("t0")}),1,0)'
        wsA[f'{col}35'].number_format = N2; wsA[f'{col}35'].font = F_N; wsA[f'{col}36'].font = F_N
    wsA['B28'] = f'=INDEX({FIRST}{ar["month"]}:{LAST}{ar["month"]},MATCH(1,{FIRST}36:{LAST}36,0))'; wsA['B28'].number_format = '0'
    wsA.cell(37, 1, 'Ticket cash received (cumulative)').font = F_N
    for i in range(1, NQ_+1):
        col = qc(i); wsA[f'{col}37'] = f"=IF({col}{ar['month']}>={R('t0')},{R('T')},0)"; wsA[f'{col}37'].font = F_N
    wsA.cell(38, 1, 'Project cumulative cash plus deal effect, ticket cash excluded').font = F_N
    for i in range(1, NQ_+1):
        col = qc(i); wsA[f'{col}38'] = f"=Project_CF!{col}{P_['cum']}+{col}{ar['jcum']}-{col}37"; wsA[f'{col}38'].font = F_N; wsA[f'{col}38'].number_format = N2
    # face multiple sensitivity & early assignment (rows 40+)
    hdr(wsA, 40, ['Face multiple sensitivity', 'Investor IRR', 'Investor MOIC'])
    for j, fx in enumerate((1.5, 1.75, 2.0, 2.25, 2.5)):
        rr = 41+j; wsA.cell(rr, 1, fx).font = F_IN; wsA.cell(rr, 1).number_format = X
        # compact closed-form flows laid out in helper rows 50+ 
        hr = 50+j; wsA.cell(hr, 1, f'Investor flow at face multiple row {rr}').font = F_NOTE
        for i in range(1, NQ_+1):
            col = qc(i); mo = f'Source_Data!{col}${S["month"]}'
            wsA[f'{col}{hr}'] = f"=IF({mo}={R('t0')},-{R('T')},0)+IF(AND({mo}>={R('handover')}+3,{mo}<{R('handover')}+3+3*{R('real_q')},MOD({mo}-({R('handover')}+3),3)=0),$A${rr}*{R('T')}*{R('appr')}*(1-{R('fric')})/{R('real_q')},0)"; wsA[f'{col}{hr}'].font = F_N; wsA[f'{col}{hr}'].number_format = N2
        wsA.cell(rr, 2, f'=(1+IRR({FIRST}{hr}:{LAST}{hr},0.05))^4-1').number_format = PC; wsA.cell(rr, 3, f'=SUMIF({FIRST}{hr}:{LAST}{hr},">0")/-SUMIF({FIRST}{hr}:{LAST}{hr},"<0")').number_format = X
    hdr(wsA, 47, ['Early assignment 24 months after funding (units at face, no appreciation)', 'Investor IRR', 'Investor MOIC'])
    for j, d_ in enumerate((0.0, 0.1, 0.2, 0.3)):
        rr = 48+j; wsA.cell(rr, 1, d_).font = F_IN; wsA.cell(rr, 1).number_format = PC; hr = 56+j
        wsA.cell(hr, 1, f'Early assignment flow row {rr}').font = F_NOTE
        for i in range(1, NQ_+1):
            col = qc(i); mo = f'Source_Data!{col}${S["month"]}'
            wsA[f'{col}{hr}'] = f"=IF({mo}={R('t0')},-{R('T')},0)+IF({mo}={R('t0')}+24,{R('face_x')}*{R('T')}*(1-$A${rr})*(1-{R('fric')}),0)"; wsA[f'{col}{hr}'].font = F_N; wsA[f'{col}{hr}'].number_format = N2
        wsA.cell(rr, 2, f'=(1+IRR({FIRST}{hr}:{LAST}{hr},0.05))^4-1').number_format = PC; wsA.cell(rr, 3, f'=SUMIF({FIRST}{hr}:{LAST}{hr},">0")/-SUMIF({FIRST}{hr}:{LAST}{hr},"<0")').number_format = X
    # fix rows overlap: face sens rows 41-45 and header at 47 ok
    # ============ Option_B ============
    wsB = sh('Option_B', f'{name}: Option B (cash), ticket EGP {T}m', 'S1 is the recommended structure. S2 and S3 are alternatives considered. Payouts after a grace period; nothing is held in a restricted account.')
    wsB.column_dimensions['A'].width = 66; wsB.column_dimensions['B'].width = 14
    for i in range(1, NQ_+1): wsB.column_dimensions[qc(i)].width = 9
    hdr(wsB, 4, ['Quarterly flows (EGP m)', 'Total'] + [''] * (NQ_+1))
    br = {'month': 5, 'slot': 6, 'net': 7, 'coll': 8, 's1': 9, 's2': 10, 's3f': 11, 's3p': 12, 's3': 13, 'i1': 14, 'i2': 15, 'i3': 16, 'j1': 17, 'j2': 18, 'j3': 19, 'cash1': 20, 'cash2': 21, 'cash3': 22,
          'c1a': 23, 'c1b': 24, 'c2a': 25, 'c2b': 26, 'c3a': 27, 'c3b': 28}
    bl = {'month': 'Month', 'slot': 'Payout window flag', 'net': 'Net cash flow before payouts (Project_CF)', 'coll': 'Collections (Project_CF)', 's1': 'S1 payouts: fixed multiple, equal instalments', 's2': 'S2 payouts: share of collections with floor and cap',
          's3f': 'S3 fixed part', 's3p': 'S3 collections share', 's3': 'S3 payouts: hybrid', 'i1': 'S1 investor cash flow', 'i2': 'S2 investor cash flow', 'i3': 'S3 investor cash flow', 'j1': 'S1 Jalour cash flow (ticket less payouts)',
          'j2': 'S2 Jalour cash flow', 'j3': 'S3 Jalour cash flow', 'cash1': 'S1 pooled cash after ticket and payouts', 'cash2': 'S2 pooled cash', 'cash3': 'S3 pooled cash', 'c1a': 'S1 coverage: net cash flow / payout', 'c1b': 'S1 coverage: cash available / payout',
          'c2a': 'S2 coverage: net cash flow / payout', 'c2b': 'S2 coverage: cash available / payout', 'c3a': 'S3 coverage: net cash flow / payout', 'c3b': 'S3 coverage: cash available / payout'}
    for key, rr in br.items(): wsB.cell(rr, 1, bl[key]).font = F_B if key in ('s1', 's2', 's3') else F_N
    eq = f"({R('B_mult')}*{R('T')}/{R('n_pay')})"
    for i in range(1, NQ_+1):
        col = qc(i); pcol = qc(i-1); mo = f'Source_Data!{col}${S["month"]}'; t0f = f"IF({mo}={R('t0')},{R('T')},0)"
        wsB[f'{col}5'] = f'={mo}'
        wsB[f'{col}6'] = f"=IF(AND({mo}>={R('first_pay')},{mo}<{R('first_pay')}+3*{R('n_pay')}),1,0)"
        wsB[f'{col}7'] = f'=Project_CF!{col}{P_["net"]}'; wsB[f'{col}8'] = f'=Project_CF!{col}{P_["coll"]}'
        wsB[f'{col}9'] = f'={col}6*{eq}'
        cumprior2 = '0' if i == 1 else f'SUM(${FIRST}$10:{pcol}10)'
        wsB[f'{col}10'] = f"=IF({mo}>={R('first_pay')},MAX(0,MIN(MAX(MIN({R('s2_pct')}*{col}8,{R('s2_cap')}*{eq}),{R('s2_floor')}*{eq}),{R('B_mult')}*{R('T')}-{cumprior2})),0)"
        wsB[f'{col}11'] = f"={col}6*{R('s3_fixed')}*{R('T')}/{R('n_pay')}"
        cumprior3 = '0' if i == 1 else f'SUM(${FIRST}$12:{pcol}12)'
        wsB[f'{col}12'] = f"=IF({col}6=1,MAX(0,MIN({R('s3_pct')}*{col}8,{R('s3_cap')}*{R('T')}-{R('s3_fixed')}*{R('T')}-{cumprior3})),0)"
        wsB[f'{col}13'] = f'={col}11+{col}12'
        for n_, (a_, j_, c_) in enumerate(((9, 14, 17), (10, 15, 18), (13, 16, 19))):
            wsB[f'{col}{j_}'] = f"=-IF({mo}={R('t0')},{R('T')},0)+{col}{a_}"
            wsB[f'{col}{c_}'] = f"=IF({mo}={R('t0')},{R('T')},0)-{col}{a_}"
            cr = 20+n_; prevc = '0' if i == 1 else f'{pcol}{cr}'
            wsB[f'{col}{cr}'] = f"={prevc}+IF({mo}={R('t0')},{R('T')},0)+{col}7-{col}{a_}"
            wsB[f'{col}{23+2*n_}'] = f'=IF({col}{a_}>0,{col}7/{col}{a_},"")'
            wsB[f'{col}{24+2*n_}'] = f'=IF({col}{a_}>0,({prevc}+IF({mo}={R("t0")},{R("T")},0)+{col}7)/{col}{a_},"")'
        for rr in range(5, 29): wsB[f'{col}{rr}'].font = F_N; wsB[f'{col}{rr}'].number_format = '0.00' if rr >= 23 else ('0' if rr in (5, 6) else N2)
    for rr in (7, 8, 9, 10, 11, 12, 13): wsB.cell(rr, 2, f'=SUM({FIRST}{rr}:{LAST}{rr})').number_format = N2
    hdr(wsB, 31, ['Structure comparison (base case)', 'S1 Fixed multiple', 'S2 % of collections', 'S3 Hybrid'])
    comp = [('Total paid (EGP m)', ['=SUM(D9:BK9)', '=SUM(D10:BK10)', '=SUM(D13:BK13)'], N1), ('Total multiple', [f"=B32/{R('T')}", f"=C32/{R('T')}", f"=D32/{R('T')}"], X),
            ('Investor IRR', [f'=(1+IRR({FIRST}14:{LAST}14,0.05))^4-1', f'=(1+IRR({FIRST}15:{LAST}15,0.05))^4-1', f'=(1+IRR({FIRST}16:{LAST}16,0.05))^4-1'], PC),
            ('Jalour cost of capital', [f'=(1+IRR({FIRST}17:{LAST}17,0.05))^4-1', f'=(1+IRR({FIRST}18:{LAST}18,0.05))^4-1', f'=(1+IRR({FIRST}19:{LAST}19,0.05))^4-1'], PC),
            ('NPV to Jalour @ discount rate', [f'=SUMPRODUCT({FIRST}17:{LAST}17,{PR("df")})', f'=SUMPRODUCT({FIRST}18:{LAST}18,{PR("df")})', f'=SUMPRODUCT({FIRST}19:{LAST}19,{PR("df")})'], N1),
            ('Lowest quarterly coverage: net cash flow / payout', [f'=MIN({FIRST}23:{LAST}23)', f'=MIN({FIRST}25:{LAST}25)', f'=MIN({FIRST}27:{LAST}27)'], '0.00'),
            ('Quarters below threshold (net cash flow basis)', [f'=COUNTIF({FIRST}23:{LAST}23,"<"&{R("gate")})', f'=COUNTIF({FIRST}25:{LAST}25,"<"&{R("gate")})', f'=COUNTIF({FIRST}27:{LAST}27,"<"&{R("gate")})'], '0'),
            ('Payout quarters', ['=COUNTIF(D9:BK9,">0")', '=COUNTIF(D10:BK10,">0")', '=COUNTIF(D13:BK13,">0")'], '0'),
            ('Lowest coverage: cash available / payout', [f'=MIN({FIRST}24:{LAST}24)', f'=MIN({FIRST}26:{LAST}26)', f'=MIN({FIRST}28:{LAST}28)'], '0.00'),
            ('Quarters below threshold (cash available basis)', [f'=COUNTIF({FIRST}24:{LAST}24,"<"&{R("gate")})', f'=COUNTIF({FIRST}26:{LAST}26,"<"&{R("gate")})', f'=COUNTIF({FIRST}28:{LAST}28,"<"&{R("gate")})'], '0'),
            ('Lowest pooled cash after ticket and payouts', [f'=MIN({FIRST}20:{LAST}20)', f'=MIN({FIRST}21:{LAST}21)', f'=MIN({FIRST}22:{LAST}22)'], N1),
            ('Last payout month', [f'=SUMPRODUCT(MAX(({FIRST}9:{LAST}9>0)*{FIRST}5:{LAST}5))', f'=SUMPRODUCT(MAX(({FIRST}10:{LAST}10>0)*{FIRST}5:{LAST}5))', f'=SUMPRODUCT(MAX(({FIRST}13:{LAST}13>0)*{FIRST}5:{LAST}5))'], '0')]
    BRES = {}
    for j, (lab, fs, fm) in enumerate(comp):
        rr = 32+j; wsB.cell(rr, 1, lab).font = F_N
        for n_, f_ in enumerate(fs):
            c = wsB.cell(rr, 2+n_, f_); c.number_format = fm; c.font = F_B; c.fill = FILL_K
        BRES[j] = rr
    wsB.cell(46, 1, 'Payback month (S1): first quarter cumulative investor cash flow is not negative').font = F_N
    wsB.cell(47, 1, 'S1 investor cumulative cash flow').font = F_N; wsB.cell(48, 1, 'Not negative (1/0)').font = F_N
    for i in range(1, NQ_+1):
        col = qc(i); wsB[f'{col}47'] = f'={col}14' if i == 1 else f'={qc(i-1)}47+{col}14'; wsB[f'{col}48'] = f'=IF(AND({col}47>=0,{col}5>{R("t0")}),1,0)'
        wsB[f'{col}47'].number_format = N2; wsB[f'{col}47'].font = F_N; wsB[f'{col}48'].font = F_N
    wsB['B46'] = f'=INDEX({FIRST}5:{LAST}5,MATCH(1,{FIRST}48:{LAST}48,0))'; wsB['B46'].number_format = '0'
    wsB.freeze_panes = 'B6'
    # ============ Summary ============
    wsM = sh('Summary', f'{name}: key outputs, ticket EGP {T}m (all cells are live links)', f'USD equivalents at {48} EGP per USD.')
    wsM.column_dimensions['A'].width = 78; wsM.column_dimensions['B'].width = 18; wsM.column_dimensions['C'].width = 18; wsM.column_dimensions['D'].width = 50
    hdr(wsM, 4, ['Metric', 'Value', 'USD m', 'Source sheet'])
    M = {}; rr = 5
    def srow(key, lab, f_, fm, usd=False, srcn=''):
        nonlocal rr
        wsM.cell(rr, 1, lab).font = F_N; c = wsM.cell(rr, 2, f_); c.number_format = fm; c.font = F_B
        if usd: wsM.cell(rr, 3, f'=B{rr}/{R("usd")}').number_format = '#,##0.0'
        wsM.cell(rr, 4, srcn).font = F_NOTE; M[key] = f'Summary!$B${rr}'; rr += 1
    def sub(t):
        nonlocal rr
        wsM.cell(rr, 1, t).font = F_B; wsM.cell(rr, 1).fill = FILL_K; rr += 1
    sub('Project (base case)')
    srow('net', 'Jalour net cash flow, nominal (EGP m)', f"={PK['net']}", N1, True, 'Project_CF'); srow('npv', 'NPV of net cash flow @ 14% (EGP m)', f"={PK['npv']}", N1, True, 'Project_CF')
    srow('peak', 'Peak cumulative cash shortage (EGP m)', f"=-{PK['peak']}", N1, True, 'Project_CF'); srow('peak_m', 'Month of peak shortage', f"={PK['peak_m']}", '0', False, 'Project_CF')
    srow('coll', 'Total collections, 80% of units sold (EGP m)', f"=Project_CF!$C${P_['coll']}", N1, True, 'Project_CF'); srow('land', 'Landlord payments (EGP m)', f"=Project_CF!$C${P_['land']}", N1, False, 'Project_CF')
    srow('cost', 'Construction, commission and SG&A (EGP m)', f"=Project_CF!$C${P_['cons']}+Project_CF!$C${P_['comm']}+Project_CF!$C${P_['sga']}", N1, False, 'Project_CF')
    sub('Funding plan')
    srow('uses', 'Total uses at peak (EGP m)', f"=Sources_Uses!B{U['tot']}", N1, True, 'Sources_Uses'); srow('bal', 'Balance of project funding (EGP m)', f"=Sources_Uses!B{U['bal']}", N1, False, 'Sources_Uses')
    srow('buf', 'Ticket in excess of modelled uses (EGP m)', f"=Sources_Uses!B{U['buf']}", N1, False, 'Sources_Uses'); srow('sh1', 'Ticket as % of modelled peak', f"=Sources_Uses!B{U['sh1']}", PC, False, 'Sources_Uses')
    srow('xs', 'Ticket above modelled peak (EGP m)', f"=Sources_Uses!B{U['xs']}", N1, False, 'Sources_Uses')
    sub('Option A (units)')
    srow('face', 'Units face value at launch list (EGP m)', "=Option_A!B5", N1, True, 'Option_A'); srow('share', 'Share of launch-tranche inventory', "=Option_A!B6", PC, False, 'Option_A'); srow('val', 'Units value at final list price (EGP m)', "=Option_A!B11", N1, False, 'Option_A')
    srow('a_irr', 'Investor IRR, base', "=Option_A!B26", PC, False, 'Option_A'); srow('a_moic', 'Investor MOIC, base', "=Option_A!B27", X, False, 'Option_A'); srow('a_pay', 'Investor payback month', "=Option_A!B28", '0', False, 'Option_A')
    srow('a_dirr', 'Investor IRR, downside', f"={SCN['Downside: sales +12 months, prices -10%']['a_irr']}", PC, False, 'Scenarios'); srow('a_dmoic', 'Investor MOIC, downside', f"={SCN['Downside: sales +12 months, prices -10%']['a_moic']}", X, False, 'Scenarios')
    srow('a_uirr', 'Investor IRR, upside', f"={SCN['Upside: prices +10%']['a_irr']}", PC, False, 'Scenarios'); srow('a_umoic', 'Investor MOIC, upside', f"={SCN['Upside: prices +10%']['a_moic']}", X, False, 'Scenarios')
    srow('j_irr', 'Jalour cost of capital', "=Option_A!B29", PC, False, 'Option_A'); srow('j_npv', 'NPV cost to Jalour @ 14% (EGP m)', "=Option_A!B30", N1, False, 'Option_A'); srow('j_nom', 'Nominal net cost to Jalour (EGP m)', "=Option_A!B31", N1, False, 'Option_A')
    sub('Option B (cash), recommended structure S1')
    srow('b_irr', 'Investor IRR, base', "=Option_B!B34", PC, False, 'Option_B'); srow('b_moic', 'Investor MOIC', "=Option_B!B33", X, False, 'Option_B'); srow('b_pay', 'Investor payback month', "=Option_B!B46", '0', False, 'Option_B')
    srow('b_j', 'Jalour cost of capital', "=Option_B!B35", PC, False, 'Option_B'); srow('b_jnpv', 'NPV cost to Jalour @ 14% (EGP m)', "=Option_B!B36", N1, False, 'Option_B')
    srow('b_cov1', 'Lowest quarterly coverage (net cash flow / payout)', "=Option_B!B37", '0.00', False, 'Option_B'); srow('b_cov1n', 'Quarters below 1.5x (net cash flow basis)', "=Option_B!B38", '0', False, 'Option_B'); srow('b_n', 'Payout quarters', "=Option_B!B39", '0', False, 'Option_B')
    srow('b_cov2', 'Lowest coverage (cash available / payout)', "=Option_B!B40", '0.00', False, 'Option_B'); srow('b_cov2n', 'Quarters below 1.5x (cash available basis)', "=Option_B!B41", '0', False, 'Option_B'); srow('b_cash', 'Lowest pooled cash after ticket and payouts (EGP m)', "=Option_B!B42", N1, False, 'Option_B')
    srow('b_dirr', 'Investor IRR if S2 (alternative)', "=Option_B!C34", PC, False, 'Option_B'); srow('b3_irr', 'Investor IRR if S3 (alternative)', "=Option_B!D34", PC, False, 'Option_B')
    sub('Scenario results: peak cumulative shortage / NPV / Option B lowest pooled cash (EGP m)')
    for nm in PD.scen_defs():
        O = SCN[nm]
        wsM.cell(rr, 1, nm).font = F_N; wsM.cell(rr, 2, f"=-{O['peak']}").number_format = N1; wsM.cell(rr, 3, f"={O['npv']}").number_format = N1; wsM.cell(rr, 4, f"={O['b_cash']}").number_format = N1
        M['scen|' + nm] = rr; rr += 1
    wsM.cell(rr, 1, '(columns: peak shortage | NPV | Option B lowest pooled cash)').font = F_NOTE; rr += 1
    sub('Sensitivities: NPV (EGP m) and peak shortage (EGP m)')
    for (fam, v), O in SEN.items():
        wsM.cell(rr, 1, f'{fam} {v}').font = F_N; wsM.cell(rr, 2, f"=-{O['peak']}").number_format = N1; wsM.cell(rr, 3, f"={O['npv']}").number_format = N1; wsM.cell(rr, 4, f"={O['a_irr']}").number_format = PC
        M[f'sens|{fam}|{v}'] = rr; rr += 1
    # ============ Reconciliation ============
    wsR = sh('Reconciliation', f'{name}: reconciliation to the source model and independent recomputation (all checks must read zero)', 'Differences are rounded to 5 decimals.')
    hdr(wsR, 4, ['Check', 'This annex', 'Reference', 'Difference', 'Reference basis'])
    checks = [('Net cash flow, nominal', f"={PK['net']}", f"={R('ref_net')}", 'Source model'), ('NPV with down payment', f"={PK['npv']}", f"={R('ref_npv')}", 'Source model'), ('Peak cumulative shortage', f"={PK['peak']}", f"={R('ref_peak')}", 'Source model'),
              ('Total collections', f"=Project_CF!$C${P_['coll']}", f"={R('ref_coll')}", 'Source model'), ('Total landlord payments', f"=Project_CF!$C${P_['land']}", f"={R('ref_land')}", 'Source model'),
              ('Construction, commission and SG&A', f"=Project_CF!$C${P_['cons']}+Project_CF!$C${P_['comm']}+Project_CF!$C${P_['sga']}", f"={R('ref_cost')}", 'Source model'),
              ('Scenario base block equals Project_CF (net)', f"={SCN['Base case']['net']}", f"={PK['net']}", 'Internal'), ('Scenario base block equals Project_CF (NPV)', f"={SCN['Base case']['npv']}", f"={PK['npv']}", 'Internal'),
              ('Option A investor total = face x appreciation x (1-friction) - ticket', "=Option_A!B17", f"=Option_A!B5*{R('appr')}*(1-{R('fric')})-{R('T')}", 'Closed form'),
              ('Option B S1 total payouts = multiple x ticket', "=Option_B!B9", f"={R('B_mult')}*{R('T')}", 'Closed form'),
              ('Total sources equal the larger of total uses and the ticket', '=Sources_Uses!B14', '=MAX(Sources_Uses!B8,Sources_Uses!B12)', 'Internal')]
    py = PD.build(k, T)
    checks += [('Python recomputation: Option A investor IRR', "=Option_A!B26", py['A']['inv']['Base']['irr'], 'Independent Python (engine2.py)'), ('Python recomputation: Option A Jalour cost of capital', "=Option_A!B29", py['A']['jal']['irr'], 'Independent Python'),
               ('Python recomputation: Option B S1 investor IRR', "=Option_B!B34", py['B']['S']['S1']['res']['Base']['irr'], 'Independent Python'), ('Python recomputation: Option B S2 investor IRR', "=Option_B!C34", py['B']['S']['S2']['res']['Base']['irr'], 'Independent Python'),
               ('Python recomputation: Option B S3 investor IRR', "=Option_B!D34", py['B']['S']['S3']['res']['Base']['irr'], 'Independent Python'), ('Python recomputation: downside peak shortage', f"=-{SCN['Downside: sales +12 months, prices -10%']['peak']}", py['fund']['downside_peak'], 'Independent Python'),
               ('Python recomputation: Option B S1 lowest pooled cash', "=Option_B!B42", py['B']['S']['S1']['res']['Base']['min_cash'], 'Independent Python')]
    for j, (lab, a_, b_, bs) in enumerate(checks):
        rr = 5+j; wsR.cell(rr, 1, lab).font = F_N; c = wsR.cell(rr, 2, a_); c.number_format = '#,##0.000000'; c.font = F_N
        c2 = wsR.cell(rr, 3, b_); c2.number_format = '#,##0.000000'; c2.font = F_IN if not (isinstance(b_, str) and b_.startswith('=')) else F_N
        c3 = wsR.cell(rr, 4, f'=ROUND(B{rr}-C{rr},5)'); c3.number_format = '0.00000'; c3.font = F_B; c3.fill = FILL_OK; wsR.cell(rr, 5, bs).font = F_NOTE
    last = 5+len(checks)-1
    wsR.cell(last+2, 1, 'Sum of absolute differences').font = F_B; wsR.cell(last+2, 4, f'=SUMPRODUCT(ABS(D5:D{last}))').number_format = '0.00000'
    wsR.cell(last+3, 1, 'Status').font = F_B; wsR.cell(last+3, 4, f'=IF(D{last+2}=0,"ALL CHECKS ZERO","CHECK FAILED")').font = F_B
    for col, w in zip('ABCDE', (70, 20, 20, 14, 40)): wsR.column_dimensions[col].width = w
    # ============ Cover ============
    wsV = wb.create_sheet('Cover', 0); wsV.sheet_view.showGridLines = False
    wsV['A1'] = f'{name} - Financial Annex'; wsV['A1'].font = Font(name=AR, size=18, bold=True, color='1F3864')
    wsV['A2'] = f'Investor ticket EGP {T} million  |  Strictly private and confidential  |  Projections only, not guarantees'; wsV['A2'].font = F_B
    notes = ['Purpose: supports the Investment Memorandum, Pitch Deck and Due Diligence Memorandum. Every figure in those documents traces to the Summary sheet.',
             'Colour code: blue = input or value transcribed from the source model; black = formula; green = link to another sheet.',
             'Units: EGP million unless stated. Quarter-end months. NPV at 14% p.a. (end-of-quarter discounting). IRRs are annualised from quarterly flows.',
             'Sheets: Inputs, Source_Data, Project_CF, Sources_Uses, Option_A, Option_B, Scenarios, Sensitivity, Summary, Reconciliation.',
             'The balance of project funding is provided by Jalour sponsor equity, project collections and other capital sources.',
             'Jalour may raise further capital at project or holding level, subject to the investor\'s stated rights.',
             'Reconciliation status:']
    for j, t_ in enumerate(notes): wsV.cell(4+j, 1, t_).font = F_N
    wsV.cell(10, 2, f"=Reconciliation!D{last+3}").font = F_B; wsV.cell(10, 2).fill = FILL_OK
    wsV.column_dimensions['A'].width = 150; wsV.column_dimensions['B'].width = 22
    order = ['Cover', 'Summary', 'Inputs', 'Source_Data', 'Project_CF', 'Sources_Uses', 'Option_A', 'Option_B', 'Scenarios', 'Sensitivity', 'Reconciliation']
    wb._sheets = [wb[n] for n in order]
    wb.save(path)
    return dict(M=M, SCN=SCN, SEN=SEN, last=last)
if __name__ == '__main__':
    info = make_annex(sys.argv[1], int(sys.argv[2]), sys.argv[3]); import json; json.dump({'M': info['M']}, open(sys.argv[3] + '.map.json', 'w'))
    print('annex written')
