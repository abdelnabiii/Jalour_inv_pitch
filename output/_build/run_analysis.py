import sys, json; sys.path.insert(0, '.')
from engine import *
TICKETS = (50, 75, 100, 125); M_B = 1.75; FACE = 2.0
R = {}
def nz(x): return None if (isinstance(x, float) and np.isnan(x)) else x
# ---------- project scenarios ----------
SC = {'Base': {}, 'Downside (sales +12m, price -10%)': dict(price=0.9, delay_q=4), 'Upside (price +10%)': dict(price=1.1),
      'Construction cost +10%': dict(cost=1.10), 'Collections slip 2 quarters': dict(slip_q=2),
      'Delivery payments at handover (stress)': dict(deliv_aligned=True),
      'Stress + downside': dict(deliv_aligned=True, price=0.9, delay_q=4)}
R['proj'] = {}
for k in P:
    R['proj'][k] = {}
    for nm, kw in SC.items():
        s = series(k, **kw); i = int(s['cum'].argmin())
        R['proj'][k][nm] = dict(net=s['net'].sum(), npv=npv(s['net']), peak=s['cum'].min(), peak_m=int(MONTHS[i]))
# ---------- funding plan ----------
PRECON = 0.05*697.16*0.30        # ASSUMPTION: 30% of the 5% professional-fee line (of construction cost) incurred before construction start
R['fund'] = {}
for k in P:
    b = series(k); pk = -b['cum'].min(); p = P[k]
    guar = pk - p['dp']                       # guarantee instalments net of collections to model peak
    uses = dict(down_payment=p['dp'], guarantee_wc_to_peak=guar, precon_assumption=PRECON)
    tot = sum(uses.values()); R['fund'][k] = dict(model_peak=pk, uses=uses, total_uses=tot, rows={})
    for T in TICKETS:
        R['fund'][k]['rows'][T] = dict(share_model_peak=T/pk, share_total_uses=min(T, tot)/tot, buffer=max(0, T-tot), excess_over_model_peak=max(0, T-pk),
                                         balance=max(0, tot-T))
# ---------- Option A ----------
R['A'] = {'inv': {}, 'jal': {}, 'pool': {}, 'face_sens': {}, 'early_exit': {}}
for k in P:
    p = P[k]
    for T in TICKETS:
        F = FACE*T
        for sc, kw in (('Base', {}), ('Downside', dict(price=0.9, delay_m=12)), ('Upside', dict(price=1.1))):
            a = optA_investor(k, T, **kw); R['A']['inv'][f'{k}|{T}|{sc}'] = dict(irr=nz(irr(a)), moic=moic(a))
        R['A']['pool'][f'{k}|{T}'] = dict(face=F, pct_retained=F/p['retained_val'], comm_pool=p['ret_comm'], adm_pool=p['ret_adm'],
                                          pct_sold=F/p['sold_total'], pct_list100=F/p['list100'])
        b = series(k)
        for src in ('retained', 'sold'):
            for ll in ('bears', 'waives'):
                d = optA_jalour(k, T, src, ll); s2 = b['net'] + d - np.eye(NQ)[p['t0']//3-1]*T   # project flows excluding the ticket cash itself
                R['A']['jal'][f'{k}|{T}|{src}|{ll}'] = dict(cost_rate=nz(irr(d, 0.0)), npv14=npv(d), nominal=d.sum(),
                      peak_ex_ticket=(b['net'] + (d - np.eye(NQ)[p['t0']//3-1]*T)).cumsum().min(), npv_project_after=npv(b['net']+d))
        # retained held longer by Jalour (cost sensitivity)
        for dly in (0, 12, 24):
            d = optA_jalour(k, T, 'retained', 'bears', retained_sale_delay_m=dly); R['A']['jal'][f'{k}|{T}|retained|bears|hold+{dly}'] = dict(cost_rate=nz(irr(d, 0.0)))
    for fx in (1.5, 1.75, 2.0, 2.25, 2.5, 2.67, 3.0):
        a = optA_investor(k, 100, face_x=fx); R['A']['face_sens'][f'{k}|{fx}'] = dict(irr=nz(irr(a)), moic=moic(a))
    for disc in (0.0, 0.10, 0.20, 0.30):                              # early assignment 24m after funding at (1-disc) x face, ex friction
        a = np.zeros(NQ); a[p['t0']//3-1] -= 100; a[(p['t0']+24)//3-1] += FACE*100*(1-disc)*(1-FRICTION)
        R['A']['early_exit'][f'{k}|{disc}'] = dict(irr=nz(irr(a)), moic=moic(a))
# ---------- Option B ----------
R['B'] = {}; R['Bsched'] = {}
for k in P:
    p = P[k]; sc = {'Base': series(k), 'Downside': series(k, price=0.9, delay_q=4), 'Upside': series(k, price=1.1), 'Stress': series(k, deliv_aligned=True)}
    rp = {'Base': retained_proceeds(k), 'Downside': retained_proceeds(k, 0.9, 12), 'Upside': retained_proceeds(k, 1.1), 'Stress': retained_proceeds(k)}
    base = sc['Base']; idx = [m//3-1 for m in pay_quarters(k)]
    for T in TICKETS:
        eq = M_B*T/12; cap, flo = 1.5*eq, 0.4*eq
        pc = calibrate_min_pct(k, T, M_B, base, cap, flo); ph = calibrate_hybrid(k, T, 1.35, M_B, 2.0, base)
        for S in ('S1', 'S2', 'S3', 'S4'):
            out = {}
            for scn, s in sc.items():
                if S == 'S1': pay = s1_fixed(k, T, M_B)
                elif S == 'S2': pay = s2_pct(k, T, M_B, pc, s, cap, flo)
                elif S == 'S3': pay = s3_hybrid(k, T, 1.35, 2.0, ph, s)
                else: pay = s4_sculpted(k, T, M_B, base['net'], rp['Base'])   # contractual schedule fixed at signing on base case
                inv = inv_flow(k, T, pay); jl = jalour_B(k, T, pay)
                c1 = cov_stats(k, pay, s['net']); c2 = cov_stats(k, pay, s['net'], rp[scn])
                cash = np.zeros(NQ); cash[p['t0']//3-1] += T; cash = (cash + s['net'] - pay).cumsum()
                cash2 = (cash + rp[scn].cumsum())
                out[scn] = dict(irr=nz(irr(inv)), moic=moic(inv), jcost=nz(irr(jl, 0.0)), paid=pay.sum(), last_pay=int(MONTHS[np.nonzero(pay)[0].max()]) if pay.sum() > 0 else None,
                                cov_min=nz(c1['min']), cov_below=c1['n_below'], cov_n=c1['n'], cov_agg=c1['agg'],
                                covr_min=nz(c2['min']), covr_below=c2['n_below'], covr_agg=c2['agg'],
                                min_cash=cash.min(), min_cash_m=int(MONTHS[cash.argmin()]), min_cash_incl_ret=cash2[:int(np.nonzero(pay)[0].max())+1].min() if pay.sum() > 0 else None,
                                npv14=npv(jl))
                if scn == 'Base': R['Bsched'][f'{k}|{T}|{S}'] = {int(MONTHS[i]): round(float(pay[i]), 3) for i in np.nonzero(pay)[0]}
            out['params'] = dict(pct=nz(pc) if S == 'S2' else (nz(ph) if S == 'S3' else None), floor=flo, cap=cap)
            R['B'][f'{k}|{T}|{S}'] = out
    R['B'][f'{k}|S1_multiple_table'] = {str(m): dict(irr=nz(irr(inv_flow(k, 100, s1_fixed(k, 100, m))))) for m in (1.5, 1.6, 1.7, 1.75, 1.8, 1.9, 2.0)}
    R['B'][f'{k}|net_by_payq'] = {int(m): round(float(base['net'][m//3-1]), 2) for m in range(pay_quarters(k)[0]-3, pay_quarters(k)[-1]+13, 3)}
# ---------- headline facts ----------
R['facts'] = {}
for k in P:
    f = D[k]['facts']; p = P[k]
    pc_launch, pa_launch = f['price_comm'][0], f['price_admin'][0]
    sc_, sa_ = np.array([x or 0 for x in f['sales_comm_pct']], float)/100, np.array([x or 0 for x in f['sales_admin_pct']], float)/100
    avg_c = f['sales_comm_80_k']/(f['area_comm']*sc_.sum()); avg_a = f['sales_admin_80_k']/(f['area_admin_sales_basis']*sa_.sum())
    R['facts'][k] = dict(launch_comm=pc_launch, launch_adm=pa_launch, final_comm=[x for x in f['price_comm'] if x][-1], final_adm=[x for x in f['price_admin'] if x][-1],
                         avg_comm=avg_c, avg_adm=avg_a, sold_comm=sc_.sum(), sold_adm=sa_.sum(), cons_total=series(k)['cons'].sum(), area_adm_sales=f['area_admin_sales_basis'])
json.dump(R, open('results.json', 'w'), indent=1, default=lambda o: float(o) if hasattr(o, '__float__') else str(o))
print('ok')
