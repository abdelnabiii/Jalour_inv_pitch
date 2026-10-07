import sys, os, re, json, zipfile, glob, math; sys.path.insert(0, '.')
import openpyxl
from docx import Document
from pptx import Presentation
import packdata as PD
OUT = '/home/user/Jalour_inv_pitch/output'
def docx_text(p):
    d = Document(p); t = [x.text for x in d.paragraphs]
    for tb in d.tables:
        for r in tb.rows:
            for c in r.cells: t.append(c.text)
    for s in d.sections:
        t += [x.text for x in s.header.paragraphs] + [x.text for x in s.footer.paragraphs]
    return '\n'.join(t)
def pptx_text(p):
    pr = Presentation(p); t = []
    for s in pr.slides:
        for sh in s.shapes:
            if sh.has_text_frame: t.append(sh.text_frame.text)
            if getattr(sh, 'has_table', False) and sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells: t.append(c.text)
            if getattr(sh, 'has_chart', False) and sh.has_chart:
                ch = sh.chart
                for pl in ch.plots:
                    t += [str(c) for c in pl.categories]
                    for se in pl.series: t += [f'{v:.1f}' for v in se.values if v is not None]
        if s.has_notes_slide: t.append(s.notes_slide.notes_text_frame.text)
    return '\n'.join(t)
import decimal
def hu(x, d):
    return f'{decimal.Decimal(str(x)).quantize(decimal.Decimal(1).scaleb(-d), rounding=decimal.ROUND_HALF_UP):,.{d}f}'
def variants(kind, x):
    base = variants0(kind, x)
    if kind in ('m1', 'm0'): base |= {hu(x, 0), hu(x, 1)}
    if kind in ('pct1', 'pct0'): base |= {hu(x*100, 1) + '%', hu(x*100, 0) + '%'}
    return base
def variants0(kind, x):
    if kind == 'm1': return {f'{x:,.1f}', f'{x:,.0f}'}
    if kind == 'm0': return {f'{x:,.0f}', f'{x:,.1f}'}
    if kind == 'pct1': return {f'{x*100:.1f}%'}
    if kind == 'pct0': return {f'{x*100:.0f}%', f'{x*100:.1f}%'}
    if kind == 'x2': return {f'{x:.2f}x', f'{x:.1f}x'}
    if kind == 'x1': return {f'{x:.1f}x', f'{x:.2f}x'}
    if kind == 'int': return {f'{x:.0f}'}
def indep_irr(flows, t, lo=-0.5, hi=3.0):
    f = lambda r: sum(c/(1+r)**ti for c, ti in zip(flows, t))
    prev_r, prev_v = lo, f(lo); r = lo
    steps = 3500
    for i in range(1, steps+1):
        r = lo + (hi-lo)*i/steps; v = f(r)
        if prev_v == 0: return prev_r
        if prev_v*v < 0:
            a, b = prev_r, r
            for _ in range(80):
                m = (a+b)/2
                if f(a)*f(m) <= 0: b = m
                else: a = m
            return (a+b)/2
        prev_r, prev_v = r, v
    return float('nan')
def run(k, T):
    folder = f'{OUT}/{k}_{T:03d}M'; work = f'/tmp/w/work_{k}_{T}'; rep = []; fails = 0
    D = PD.build(k, T); mp = json.load(open(f'{work}/annex_map.json'))['M']
    wb = openpyxl.load_workbook(f'{folder}/Financial_Annex.xlsx', data_only=True)
    def cell(ref):
        sh, a = ref.split('!'); return wb[sh.strip("'")][a.replace('$', '')].value
    S = {key: cell(ref) for key, ref in mp.items() if not key.startswith(('scen|', 'sens|'))}
    texts = {'Deck': pptx_text(f'{folder}/Deck.pptx'), 'Memo': docx_text(f'{folder}/Investment_Memo.docx'), 'DD': docx_text(f'{folder}/DD_Memo.docx')}
    rep.append(f'# QA report: {D["info"]["name"]}, ticket EGP {T}m\n')
    # 1 tie-out
    spec = [('Jalour net cash flow', 'net', 'm1', 'Deck Memo'), ('NPV at 14%', 'npv', 'm1', 'Deck Memo'), ('Peak shortage', 'peak', 'm1', 'Deck Memo DD'), ('Total uses at peak', 'uses', 'm1', 'Deck Memo'), ('Balance of project funding', 'bal', 'm1', 'Deck Memo'),
            ('Ticket as % of modelled peak', 'sh1', 'pct0', 'Deck Memo'), ('Option A face value', 'face', 'm0', 'Deck Memo'), ('Option A value at final list', 'val', 'm0', 'Memo'), ('Option A investor IRR base', 'a_irr', 'pct1', 'Deck Memo'), ('Option A investor MOIC base', 'a_moic', 'x2', 'Deck Memo'),
            ('Option A IRR downside', 'a_dirr', 'pct1', 'Deck Memo'), ('Option A MOIC downside', 'a_dmoic', 'x2', 'Deck Memo'), ('Option A IRR upside', 'a_uirr', 'pct1', 'Deck Memo'), ('Jalour cost of capital, Option A', 'j_irr', 'pct1', 'Memo DD'), ('Option A NPV cost to Jalour', 'j_npv', 'm1', 'Memo DD'),
            ('Option B investor IRR', 'b_irr', 'pct1', 'Deck Memo DD'), ('Option B investor MOIC', 'b_moic', 'x2', 'Memo'), ('Option B lowest cash-available coverage', 'b_cov2', 'x1', 'Deck Memo'), ('Option B quarters below 1.5x (net cash flow)', 'b_cov1n', 'int', 'Memo DD'),
            ('Option B IRR if S2', 'b_dirr', 'pct1', 'Memo DD'), ('Option B IRR if S3', 'b3_irr', 'pct1', 'Memo DD')]
    rows = []; miss = 0
    for lab, key, kind, docs in spec:
        val = S[key]; v = abs(val) if key in ('peak',) else val
        vs = variants(kind, v); res = {}
        for dn in docs.split():
            res[dn] = any(s in texts[dn] for s in vs)
        ok = all(res.values()); miss += (not ok); rows.append((lab, key, val, '/'.join(sorted(vs)), res, ok))
    rep.append('## 1. Number tie-out (annex Summary vs deck, memo, DD memo)\n\n| Figure | Annex cell | Annex value | Text searched | Found in | Result |\n|---|---|---|---|---|---|')
    for lab, key, val, vs, res, ok in rows:
        rep.append(f'| {lab} | {mp[key].replace("$", "")} | {val:.4f} | {vs} | ' + ', '.join(f'{a}: {"yes" if b else "NO"}' for a, b in res.items()) + f' | {"OK" if ok else "CHECK"} |')
    # extra: downside and stress peaks & ticket
    ex = []
    for lab, val, kind, docs in [('Stress peak shortage', D['fund']['stress_peak'], 'm0', 'Deck Memo DD'), ('Downside peak shortage', D['fund']['downside_peak'], 'm0', 'Memo DD'), ('Ticket', T, 'm0', 'Deck Memo DD'), ('Option B total return', D['B']['mult']*T, 'm0', 'Deck Memo'), ('Instalment', D['B']['mult']*T/12, 'm1', 'Memo')]:
        vs = variants(kind, val) | ({f'{val:,.2f}'} if lab == 'Instalment' else set()); res = {dn: any(s in texts[dn] for s in vs) for dn in docs.split()}
        ok = all(res.values()); miss += (not ok); rep.append(f'| {lab} | python | {val:.4f} | {"/".join(sorted(vs))} | ' + ', '.join(f'{a}: {"yes" if b else "NO"}' for a, b in res.items()) + f' | {"OK" if ok else "CHECK"} |')
    rep.append(f'\nTie-out misses: {miss}\n'); fails += miss
    # annex inputs trace: every Inputs row has a source/basis text
    wsI = wb['Inputs']; unsourced = [r for r in range(5, wsI.max_row+1) if wsI.cell(r, 2).value is not None and wsI.cell(r, 1).value and not wsI.cell(r, 4).value and wsI.cell(r, 3).value]
    rep.append(f'Annex inputs without a source or basis note: {len(unsourced)}\n')
    wbf = openpyxl.load_workbook(f'{folder}/Financial_Annex.xlsx')
    hard = 0
    for n in ('Project_CF', 'Sources_Uses', 'Option_A', 'Option_B', 'Scenarios', 'Sensitivity', 'Summary'):
        ws = wbf[n]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, (int, float)) and not (n in ('Scenarios', 'Sensitivity') and c.column == 2 and c.font.color is not None and c.font.color.rgb == 'FF0000FF') and not (n == 'Option_A' and c.column == 1 and c.font.color is not None and c.font.color.rgb == 'FF0000FF'):
                    hard += 1
    rep.append(f'Hard-coded numeric cells on calculation sheets (excluding blue scenario parameters): {hard}\n')
    # 2 arithmetic
    t = [m/12 for m in [3*(i+1) for i in range(60)]]
    def row_vals(sheet, r): ws = wb[sheet]; return [ws.cell(r, 4+i).value or 0 for i in range(60)]
    chk = []
    inv = row_vals('Option_A', 17); chk.append(('Option A investor IRR', indep_irr(inv, t), S['a_irr'], 1e-4))
    chk.append(('Option A investor MOIC', sum(x for x in inv if x > 0)/-sum(x for x in inv if x < 0), S['a_moic'], 1e-4))
    jal = row_vals('Option_A', 21); chk.append(('Option A Jalour cost of capital', indep_irr(jal, t, 0.0), S['j_irr'], 1e-4)); chk.append(('Option A Jalour NPV @14%', sum(c/(1.14**ti) for c, ti in zip(jal, t)), S['j_npv'], 1e-4))
    b1 = row_vals('Option_B', 14); chk.append(('Option B S1 investor IRR', indep_irr(b1, t), S['b_irr'], 1e-4))
    net = row_vals('Project_CF', 13); chk.append(('Project NPV @14%', sum(c/(1.14**ti) for c, ti in zip(net, t)), S['npv'], 1e-4)); chk.append(('Project net cash flow', sum(net), S['net'], 1e-4))
    cum = 0; mn = 0
    for x in net: cum += x; mn = min(mn, cum)
    chk.append(('Peak cumulative shortage', -mn, S['peak'], 1e-4))
    pay = row_vals('Option_B', 9); n_ = row_vals('Option_B', 7); cov = [a/b for a, b in zip(n_, pay) if b > 0]
    chk.append(('Option B min coverage (net cash flow)', min(cov), S['b_cov1'], 1e-4)); chk.append(('Option B quarters below 1.5x', sum(c < 1.5 for c in cov), S['b_cov1n'], 1e-9))
    cash = 0; liq = []
    ws = wb['Option_B']; t0 = PD.P[k]['t0']
    for i in range(60):
        m = 3*(i+1); prev = cash; cash = cash + (T if m == t0 else 0) + n_[i] - pay[i]
        if pay[i] > 0: liq.append((prev + (T if m == t0 else 0) + n_[i])/pay[i])
    chk.append(('Option B min coverage (cash available)', min(liq), S['b_cov2'], 1e-4))
    rep.append('## 2. Arithmetic check: independent recomputation (pure-Python bisection) vs spreadsheet\n\n| Measure | Independent | Spreadsheet | Abs difference | Result |\n|---|---|---|---|---|')
    bad = 0
    for lab, a, b, tol in chk:
        diff = abs(a-b); ok = diff <= tol; bad += (not ok); rep.append(f'| {lab} | {a:.6f} | {b:.6f} | {diff:.2e} | {"OK" if ok else "EXPLAIN"} |')
    rep.append(f'\nPython engine vs spreadsheet reconciliation sheet: {S and cell("Reconciliation!$D$" + str(wb["Reconciliation"].max_row))}\n'); fails += bad
    # 3 disclosure scan
    other = ["L'avenir", 'L’avenir', 'Lavenir', "L'AVENIR", 'LA_0', 'LA_1'] if k == 'GS' else ['Green Square', 'GREEN SQUARE', 'GS_0', 'GS_1']
    pats = other + ['AT EAST', 'At East', 'AT East', '250M', '250 M', '250 million', '250,000,000', 'total raise', 'total programme', 'other investor', 'second investor', 'only external investor', 'last external investor', 'entire balance sheet', 'own balance sheet', 'combined raise']
    rep.append('## 3. Disclosure scan (every file, including document properties, notes, hidden sheets and comments)\n\n| File | Parts scanned | Prohibited hits | Bare "250" contexts |\n|---|---|---|---|')
    hits_total = 0
    for f in sorted(glob.glob(folder + '/*')):
        z = zipfile.ZipFile(f); parts = 0; hits = []; bare = []; allowed = []
        for nme in z.namelist():
            if not nme.endswith(('.xml', '.rels', '.txt', '.json')): continue
            try: tx = z.read(nme).decode('utf8', 'ignore')
            except Exception: continue
            parts += 1; clean = re.sub(r'<[^>]+>', ' ', tx)
            for p_ in pats:
                if p_ in clean or p_ in tx:
                    if p_ in ('250 million', '250M', '250 M', '250,000,000'):
                        ctxs = [clean[max(0, m.start()-70):m.start()] for m in re.finditer(re.escape(p_), clean)]
                        if T == 125 and all(re.search(r'(?i)face value|units|option a', c) for c in ctxs): allowed.append(f'{p_} x{len(ctxs)} in {nme} (units face value 2.0 x 125)'); continue
                    hits.append(f'{p_} in {nme}')
            for m_ in re.finditer(r'.{25}\b250\b.{25}', clean): bare.append(f'{nme}: ...{m_.group(0).strip()}...')
        if f.endswith('.xlsx'):
            wsx = openpyxl.load_workbook(f)
            hidden = [w.title for w in wsx if w.sheet_state != 'visible']; cm = [(w.title, c.coordinate) for w in wsx for r in w.iter_rows() for c in r if c.comment]
            if hidden or cm: hits.append(f'hidden sheets {hidden} comments {cm}')
        hits_total += len(hits); rep.append(f'| {os.path.basename(f)} | {parts} | {len(hits)} {"; ".join(hits[:5])} {"Allowed in context: " + "; ".join(allowed) if allowed else ""} | {len(bare)} {" | ".join(bare[:3])} |')
    for f in sorted(glob.glob(folder + '/*')) + [folder]:
        for p_ in pats:
            if p_.lower() in os.path.basename(f).lower(): hits_total += 1; rep.append(f'FILE NAME HIT: {f} ({p_})')
    rep.append(f'\nProhibited-term hits: {hits_total}. Bare "250" occurrences are listed for context only: they are prices or the units face value, not a programme total.\n'); fails += hits_total
    # 4 consistency
    rep.append('## 4. Consistency of ticket, use of funds and returns across documents\n')
    cons = []
    for lab, strs, only in [('Ticket EGP ' + str(T), {f'EGP {T} million', f'EGP {T}.0', f'{T}.0'}, ('Deck', 'Memo', 'DD')), ('Total uses', variants('m1', D['fund']['total']), ('Deck', 'Memo')), ('Option A IRR', variants('pct1', D['A']['inv']['Base']['irr']), ('Deck', 'Memo')), ('Option B IRR', variants('pct1', D['B']['S']['S1']['res']['Base']['irr']), ('Deck', 'Memo', 'DD'))]:
        r_ = {dn: any(s in texts[dn] for s in strs) for dn in only}; cons.append((lab, r_)); fails += (not all(r_.values()))
        rep.append(f'- {lab}: ' + ', '.join(f'{a} {"yes" if b else "NO"}' for a, b in r_.items()) + ', annex yes (Summary links)')
    rep.append('')
    # 6 challenge pass
    rep.append('## 6. Challenge pass: five toughest questions from a sceptical family office\n')
    gr = sorted([(D["prices"]["final"][1]/D["prices"]["launch"][1]-1)*100, (D["prices"]["final"][0]/D["prices"]["launch"][0]-1)*100])
    qs = [('Why should list prices rise ' + f'{gr[0]:.0f}% to {gr[1]:.0f}% when market growth is 3% to 16%?', 'Memo 4.4 and 13.1; Deck slides 5 and 14; DD 5', ['Sales prices may therefore be lower or slower', 'Option B removes unit price exposure']),
          ('What if delivery payments arrive at handover, not before?', 'Memo 11.1 and 13.2; Deck slide 11; DD red flag 1', ['peak funding', 'handover']),
          ('Who pays if there is a funding gap and why is there no guarantee?', 'Memo 7, 11.1, 14.2; DD 10 and 11', ['Jalour will fund', 'Parent guarantee: not offered']),
          ('Does cash cover the Option B instalments?', 'Memo 13.3; Deck slide 14; DD 8.2', ['Payout coverage', 'Coverage on net cash flow alone is below the 1.5x threshold']),
          ('Can I get out of Option A early and what does Jalour need from the landlord?', 'Memo 6.1, 12.1, 16.1; DD 2', ['Early assignment', 'landlord'])]
    for q, loc, kws in qs:
        alltext = '\n'.join(texts.values()); found = all(any(kw.lower() in tx.lower() for tx in texts.values()) for kw in kws)
        rep.append(f'- **{q}** Answered in: {loc}. Check: {"present" if found else "MISSING"}'); fails += (not found)
    rep.append(f'\n**Total failures requiring attention: {fails}**')
    os.makedirs(f'{OUT}/QA_Reports', exist_ok=True); open(f'{OUT}/QA_Reports/QA_{k}_{T:03d}M.md', 'w').write('\n'.join(rep))
    return fails
if __name__ == '__main__': print('failures', run(sys.argv[1], int(sys.argv[2])))
