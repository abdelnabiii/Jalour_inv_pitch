"""Single source of numbers for every pack file. build(k, T) -> dict. All EGP million unless stated."""
import sys, json, copy
sys.path.insert(0, __file__.rsplit('/', 1)[0])
from engine2 import *
USD = 48.0
INFO = {
 'GS': dict(short='GS', name='Green Square', land=14470, footprint=4341, retail=4341, office=8682, bua=13023, admin_sales_area=10129, height='G+2',
            sales_start=7, con_start=13, con_months=36, delivery=48, guarantee=800.0, dp=80.0, sched=[45, 90, 135, 180, 180, 90], excess_years=8,
            total100=2816.22375, total80=2252.979, inflow=1267.3006875, cons=697.16, comm=90.11916, sga=104.574, retained=0.20),
 'LA': dict(short='LA', name="L'avenir", land=14910, footprint=4473, retail=4473, office=8946, bua=13419, admin_sales_area=10437, height='G+2',
            sales_start=13, con_start=19, con_months=36, delivery=54, guarantee=850.0, dp=85.0, sched=[0, 90, 180, 225, 180, 90], excess_years=8,
            total100=3342.1696875, total80=2673.73575, inflow=1503.976359375, cons=697.16, comm=106.94943, sga=104.574, retained=0.20)}
MODEL_REF = {'GS': dict(net=375.4475275, npv=215.31139254892392, peak=-102.5, land=985.6783125, coll=2252.979, cost=891.85316),
             'LA': dict(net=595.292929375, npv=307.5323710915294, peak=-102.0180625, land=1169.759390625, coll=2673.73575, cost=908.68343)}
def liq_cash(k, T, pay, s):
    c = np.zeros(NQ); c[P[k]['t0']//3-1] += T; return (c + s['net'] - pay).cumsum()
def liq_cov(k, T, pay, s):
    cash = liq_cash(k, T, pay, s); out = []
    for m in pay_quarters(k):
        i = m//3-1
        if pay[i] > 1e-9: out.append((m, (cash[i-1] + s['net'][i])/pay[i] if i else float('nan')))
    return out
def scen_defs(): return {
 'Base case': {}, 'Downside: sales +12 months, prices -10%': dict(price=0.9, delay_q=4), 'Upside: prices +10%': dict(price=1.1),
 'Construction cost +10%': dict(cost=1.10), 'Collections slip 2 quarters': dict(slip_q=2),
 'Delivery payments at handover': dict(deliv_aligned=True), 'Handover payments and downside combined': dict(deliv_aligned=True, price=0.9, delay_q=4)}
def build(k, T):
    p = P[k]; I = INFO[k]; Lh = launch(k); b = series(k); F = FACE_X*T
    D = dict(k=k, T=T, info=I, folder=f'{k}_{T:03d}M')
    D['proj'] = {nm: dict(net=float(series(k, **kw)['net'].sum()), npv=npv(series(k, **kw)['net']), peak=float(series(k, **kw)['cum'].min()), peak_m=int(MONTHS[series(k, **kw)['cum'].argmin()])) for nm, kw in scen_defs().items()}
    D['base'] = dict(net=float(b['net'].sum()), npv=npv(b['net']), peak=float(b['cum'].min()), peak_m=int(MONTHS[b['cum'].argmin()]), coll=float(b['coll'].sum()), land=float(b['land'].sum()),
                     cost=float((b['cons']+b['comm']+b['sga']).sum()), cons=float(b['cons'].sum()), commission=float(b['comm'].sum()), sga=float(b['sga'].sum()),
                     npv_nodp=None)
    # funding plan
    pk = -b['cum'].min(); precon = 0.05*697.16*0.30
    uses = dict(dp=I['dp'], guar_wc=pk-I['dp'], precon=precon); tot = sum(uses.values())
    D['fund'] = dict(model_peak=pk, uses=uses, total=tot, ticket=T, applied=min(T, tot), balance=max(0, tot-T), excess_over_uses=max(0, T-tot), excess_over_peak=max(0, T-pk), share_peak=T/pk, share_uses=min(T, tot)/tot,
                     stress_peak=-D['proj']['Delivery payments at handover']['peak'], downside_peak=-D['proj']['Downside: sales +12 months, prices -10%']['peak'])
    # Option A
    a = optA_inv(k, T); d, f = optA_jal(k, T, 'launch'); dp_, fp = optA_jal(k, T, 'prorata')
    appr = appreciation(k)
    A = dict(face=F, face_x=FACE_X, appr=appr, value_at_handover=F*appr, launch_prices=Lh['p0'], final_prices=Lh['fin'], launch_value=Lh['Vc']+Lh['Va'], launch_months=Lh['months'],
             retail_share=Lh['Vc']/(Lh['Vc']+Lh['Va']), f=f, tranche_pct=f, retail_face=F*Lh['Vc']/(Lh['Vc']+Lh['Va']), office_face=F*Lh['Va']/(Lh['Vc']+Lh['Va']),
             area_retail=F*Lh['Vc']/(Lh['Vc']+Lh['Va'])*1000/Lh['p0'][0], area_office=F*Lh['Va']/(Lh['Vc']+Lh['Va'])*1000/Lh['p0'][1])
    A['inv'] = {}
    for nm, kw in (('Base', {}), ('Downside', dict(price=0.9, delay_m=12)), ('Upside', dict(price=1.1)), ('Appreciation 40%', dict(g=0.40)), ('Appreciation 60%', dict(g=0.60)), ('Flat prices (0%)', dict(g=0.0))):
        x = optA_inv(k, T, **kw); A['inv'][nm] = dict(irr=irr(x), moic=moic(x), flows=x.tolist())
    A['jal'] = dict(irr=irr(d, 0.0), npv=npv(d), nominal=float(d.sum()), flows=d.tolist(), peak_ex_ticket=float((b['net']+d-np.eye(NQ)[p['t0']//3-1]*T).cumsum().min()),
                    npv_project_after=npv(b['net']+d), landlord_on_units=LL*F, prorata_irr=irr(dp_, 0.0), prorata_npv=npv(dp_))
    A['sens_face'] = {fx: dict(irr=irr(optA_inv(k, T, face_x=fx)), moic=moic(optA_inv(k, T, face_x=fx)), jirr=irr(optA_jal(k, T, 'launch', face_x=fx)[0], 0.0)) for fx in (1.5, 1.75, 2.0, 2.25, 2.5)}
    A['early'] = {}
    for dsc in (0.0, 0.1, 0.2, 0.3):
        x = np.zeros(NQ); x[p['t0']//3-1] -= T; x[(p['t0']+24)//3-1] += F*appr*0.85*(1-dsc)*(1-FRICTION)  # early assignment at month 24: units at 85% of launch-to-handover appreciation path is an ASSUMPTION; use plain face
        x = np.zeros(NQ); x[p['t0']//3-1] -= T; x[(p['t0']+24)//3-1] += F*(1-dsc)*(1-FRICTION); A['early'][dsc] = dict(irr=irr(x), moic=moic(x))
    D['A'] = A
    # Option B
    Bm = M_B2; eq = Bm*T/12; cap, flo = 1.5*eq, 0.4*eq
    pc = calibrate_min_pct(k, T, Bm, b, cap, flo); ph = calibrate_hybrid(k, T, 1.8, Bm, 2.8, b)
    scs = {'Base': b, 'Downside': series(k, price=0.9, delay_q=4), 'Upside': series(k, price=1.1), 'Stress': series(k, deliv_aligned=True)}
    B = dict(mult=Bm, grace=GRACE_M, n=PAY_Q, pay_months=pay_quarters(k), params=dict(S2_pct=pc, S2_floor=flo, S2_cap=cap, S3_pct=ph, S3_fixed=1.8, S3_cap=2.8), S={})
    for S, nm in (('S1', 'Fixed multiple, equal quarterly instalments'), ('S2', 'Percentage of collections with floor and cap'), ('S3', 'Hybrid: fixed minimum plus collections share')):
        res = {}
        for sn, s in scs.items():
            if S == 'S1': pay = s1_fixed(k, T, Bm)
            elif S == 'S2': pay = s2_pct(k, T, Bm, pc, s, cap, flo)
            else: pay = s3_hybrid(k, T, 1.8, 2.8, ph, s)
            inv = inv_flow(k, T, pay); jl = jalour_B(k, T, pay); cs = cov_stats(k, pay, s['net']); lc = liq_cov(k, T, pay, s); cash = liq_cash(k, T, pay, s)
            res[sn] = dict(pay=pay.tolist(), irr=irr(inv), moic=moic(inv), jirr=irr(jl, 0.0), jnpv=npv(jl), strict_min=cs['min'], strict_below=cs['n_below'], n=cs['n'], strict_agg=cs['agg'],
                           liq_min=min(c for _, c in lc) if lc else float('nan'), liq_below=sum(c < 1.5 for _, c in lc), liq=lc, min_cash=float(cash.min()), min_cash_m=int(MONTHS[cash.argmin()]),
                           last=int(MONTHS[np.nonzero(pay)[0].max()]), paid=float(pay.sum()))
        B['S'][S] = dict(name=nm, res=res)
    D['B'] = B
    def payback(flows):
        c = np.cumsum(np.array(flows)); i0 = p['t0']//3-1
        for i in range(i0+1, NQ):
            if c[i] >= -1e-9: return int(MONTHS[i])
        return None
    D['A_payback'] = payback(A['inv']['Base']['flows']); D['B_payback'] = payback(inv_flow(k, T, np.array(B['S']['S1']['res']['Base']['pay'])))
    # break-evens (S1, no Jalour support): price factor / delay at which pooled SPV cash after payouts first goes negative after t0+grace
    def min_cash_after(price=1.0, delay_q=0):
        s = series(k, price=price, delay_q=delay_q); pay = s1_fixed(k, T, Bm); cash = liq_cash(k, T, pay, s); i0 = p['t0']//3-1+8
        return cash[i0:].min()
    lo, hi = 0.3, 1.0
    D['breakeven'] = dict(price_for_cash=None)
    # investor Option A break-even resale price vs launch list
    D['breakeven']['A_price_factor_moic1'] = T/(F*appr*(1-FRICTION))
    D['sens'] = {'price': {f: dict(npv=npv(series(k, price=f)['net']), peak=float(series(k, price=f)['cum'].min()), net=float(series(k, price=f)['net'].sum())) for f in (0.8, 0.9, 1.0, 1.1, 1.2)},
                 'delay': {q: dict(npv=npv(series(k, delay_q=q)['net']), peak=float(series(k, delay_q=q)['cum'].min()), net=float(series(k, delay_q=q)['net'].sum())) for q in (0, 2, 4, 6, 8)},
                 'cost': {f: dict(npv=npv(series(k, cost=f)['net']), peak=float(series(k, cost=f)['cum'].min()), net=float(series(k, cost=f)['net'].sum())) for f in (1.0, 1.1, 1.2)},
                 'slip': {q: dict(npv=npv(series(k, slip_q=q)['net']), peak=float(series(k, slip_q=q)['cum'].min()), net=float(series(k, slip_q=q)['net'].sum())) for q in (0, 1, 2, 4)}}
    D['series'] = {a_: b[a_].tolist() for a_ in ('coll', 'land', 'cons', 'comm', 'sga', 'net', 'cum')}
    D['series']['months'] = MONTHS.tolist()
    # annual aggregation (project years 1..11) of collections / landlord / costs
    yrs = {}
    for y in range(1, 12):
        sel = [(3*y*4//4 - 11 + i) for i in ()]  # placeholder, replaced below
    ann = []
    for y in range(1, 12):
        idx = [i for i, m in enumerate(MONTHS) if 12*(y-1) < m <= 12*y]
        ann.append(dict(year=y, coll=float(sum(b['coll'][i] for i in idx)), land=float(sum(b['land'][i] for i in idx)), cost=float(sum(b['cons'][i]+b['comm'][i]+b['sga'][i] for i in idx)),
                        net=float(sum(b['net'][i] for i in idx)), cum=float(b['cum'][idx[-1]])))
    D['annual'] = ann
    # tranche table
    D['tranches'] = [dict(month=t['month'], comm=t['sales_comm']/1000, admin=t['sales_admin']/1000, pc=t['price_comm'], pa=t['price_admin']) for t in TR[k]['tranches']]
    ach = TR[k]['tranches']; area_c = sum(t['area_comm'] for t in ach); area_a = sum(t['area_admin'] for t in ach)
    D['prices'] = dict(launch=Lh['p0'], final=Lh['fin'], avg=(sum(t['sales_comm'] for t in ach)/area_c, sum(t['sales_admin'] for t in ach)/area_a), sold_area=(area_c, area_a))
    yrs = (I['delivery']+2)/12; mid = 113.0       # ASSUMPTION: model month 1 = December 2026; office asking mid-point EGP 113,000 per m2 (106,000 to 120,000)
    pl_c_, pl_a_ = Lh['p0']; pf_c_, pf_a_ = Lh['fin']; ap_c_, ap_a_ = D['prices']['avg']
    D['pricecmp'] = dict(years=yrs, ask_mid_off=mid, off_vs_ask=pf_a_/mid-1, off_cagr=(pf_a_/mid)**(1/yrs)-1, launch_off_vs_ask=pl_a_/mid-1, retail_g=pf_c_/pl_c_-1, office_g=pf_a_/pl_a_-1, retail_avg_g=ap_c_/pl_c_-1, office_avg_g=ap_a_/pl_a_-1, blended=A['appr']-1)

    if T == 125:
        ok = 'LA' if k == 'GS' else 'GS'; Io = INFO[ok]; bo = series(ok); pk_o = float(bo['cum'].min())
        comb = {}
        for nm, kw in scen_defs().items():
            ss, so = series(k, **kw), series(ok, **kw); cum = ss['cum'] + so['cum']; i = int(cum.argmin())
            comb[nm] = dict(self_peak=float(ss['cum'].min()), other_peak=float(so['cum'].min()), comb_peak=float(cum.min()), comb_peak_m=int(MONTHS[i]), net=float((ss['net']+so['net']).sum()), npv=npv(ss['net']+so['net']), self_at_peak=float(ss['cum'][i]), other_at_peak=float(so['cum'][i]),
                            other_npv=npv(so['net']), other_net=float(so['net'].sum()))
        cc = b['cum'] + bo['cum']; i = int(cc.argmin()); dn = series(k, price=0.9, delay_q=4); jd = int(dn['cum'].argmin())
        D['combined'] = dict(other_k=ok, other=dict(Io, k=ok), months=MONTHS[:43].tolist(), self_cum=b['cum'][:43].tolist(), other_cum=bo['cum'][:43].tolist(), comb_cum=cc[:43].tolist(), comb_min=float(cc.min()), comb_min_m=int(MONTHS[i]),
                             other_net_series=bo['net'].tolist(), other_base=dict(coll=float(bo['coll'].sum()), land=float(bo['land'].sum()), cost=float((bo['cons']+bo['comm']+bo['sga']).sum()), net=float(bo['net'].sum()), npv=npv(bo['net']), peak=pk_o, peak_m=int(MONTHS[bo['cum'].argmin()])),
                             scen=comb, own_dn_peak_m=int(MONTHS[jd]), other_base_cum_at_own_dn_peak=float(bo['cum'][jd]), comb_pos_month=int(MONTHS[[j for j in range(len(cc)) if cc[j] > 0 and j >= i][0]]),
                             other_scen_series={nm: series(ok, **kw)['net'].tolist() for nm, kw in scen_defs().items()})
    return D
if __name__ == '__main__':
    D = build('GS', 100)
    print(json.dumps({a_: D[a_] for a_ in ('base', 'fund')}, indent=1, default=float))
    print('A', {a_: (D['A'][a_] if not isinstance(D['A'][a_], dict) else None) for a_ in D['A']})
    print({n: (round(v['irr'], 4), round(v['moic'], 3)) for n, v in D['A']['inv'].items()}, D['A']['jal']['irr'], D['A']['jal']['npv'])
    for S, v in D['B']['S'].items(): print(S, {sn: (round(r['irr'], 4), r['strict_below'], round(r['liq_min'], 2), round(r['min_cash'], 1)) for sn, r in v['res'].items()})
