"""Structuring engine: quarterly project cash flows rebuilt from the source model + investor option maths.
All EGP million. Time axis: quarter-end months 3,6,...; t(years)=month/12; discount rate 14% p.a. (model, 'Green Square'!C72)."""
import json, numpy as np
from scipy.optimize import brentq
D = json.load(open(__file__.replace('engine.py', 'model_series.json')))
NQ = 60                                          # extended horizon (M180) so delays do not truncate
MONTHS = np.array([3*(i+1) for i in range(NQ)])
T_Y = MONTHS/12
DISC = 0.14
def pad(a): a = np.array(a, float); return np.concatenate([a, np.zeros(NQ-len(a))])
P = {   # project parameters taken from the model
 'GS': dict(name='Green Square', delivery=48, t0=3,  dp=80.0, minG=800.0, sold_total=2252.979, list100=2816.22375, retained_val=563.24475, ret_comm=1163.930625*0.2, ret_adm=1652.29312*0.2, sched_q=24),
 'LA': dict(name="L'avenir",     delivery=54, t0=12, dp=85.0, minG=850.0, sold_total=2673.73575, list100=3342.1696875, retained_val=668.43393750, ret_comm=1346.0934375*0.2, ret_adm=1996.0762*0.2, sched_q=24)}
COMM = 0.04; LL = 0.35
def series(k, price=1.0, delay_q=0, cost=1.0, slip_q=0, deliv_aligned=False, pre_con=0.0):
    x = {a: pad(b) for a, b in D[k].items() if a != 'facts'}
    p = P[k]
    coll = x['collections'].copy()
    if deliv_aligned:
        dq = list(MONTHS).index(p['delivery'])
        coll = coll - x['delivery_pay']; coll[dq] += x['delivery_pay'].sum()
    shift = delay_q + slip_q
    coll = np.roll(coll, shift) if shift else coll
    if shift: coll[:shift] = 0
    comm = np.roll(x['commission'], delay_q) if delay_q else x['commission'].copy()
    if delay_q: comm[:delay_q] = 0
    coll = coll*price; comm = comm*price
    cons = x['construction']*cost; sga = x['sga']*cost
    L = x['landlord'].copy(); sched = L.copy(); sched[p['sched_q']:] = 0; exc = L - sched
    exc_tot_base = exc.sum()
    exc_tot = max(0.0, LL*p['list100']*price - p['minG'])
    exc = exc*(exc_tot/exc_tot_base) if exc_tot_base else exc
    land = sched + exc
    pc = np.zeros(NQ)
    if pre_con: pc[list(MONTHS).index(12)] = pre_con      # pre-construction spend ahead of construction start (assumption)
    net = coll - land - cons - comm - sga - pc
    return dict(coll=coll, land=land, sched=sched, exc=exc, cons=cons, comm=comm, sga=sga, net=net, cum=net.cumsum(), pc=pc)
def npv(cf, r=DISC): return float((cf/(1+r)**T_Y).sum())
def irr(cf, lo=-0.3):
    """Smallest root of NPV(r)=0 on a -30%..300% grid (handles sign-changing flows)."""
    cf = np.asarray(cf, float)
    if not (cf.min() < 0 < cf.max()): return float('nan')
    f = lambda r: (cf/(1+r)**T_Y).sum()
    grid = np.linspace(lo, 3.0, int((3.0-lo)*200)+1); vals = [f(r) for r in grid]
    for i in range(len(grid)-1):
        if vals[i] == 0: return grid[i]
        if vals[i]*vals[i+1] < 0: return brentq(f, grid[i], grid[i+1])
    return float('nan')
def vec(d):                                       # {month: amount} -> array
    a = np.zeros(NQ)
    for m, v in d.items(): a[m//3-1] += v
    return a
# ---------------- Option A ----------------
FRICTION = 0.03        # ASSUMPTION: investor resale cost (broker + developer transfer fee)
REAL_Q = 4             # ASSUMPTION: units resold evenly over 4 quarters starting 1 quarter after handover
def optA_investor(k, T, face_x=2.0, price=1.0, delay_m=0, friction=FRICTION):
    p = P[k]; F = face_x*T; a = np.zeros(NQ); a[p['t0']//3-1] -= T
    start = p['delivery'] + 3 + delay_m
    for i in range(REAL_Q): a[(start+3*i)//3-1] += F*price*(1-friction)/REAL_Q
    return a
def moic(cf): return cf[cf>0].sum()/-cf[cf<0].sum()
def optA_jalour(k, T, source='retained', landlord='bears', face_x=2.0, price=1.0, delay_m=0, retained_sale_delay_m=0):
    """Delta cash flow to Jalour (+T at t0 less forgone value); returns array of incremental flows."""
    p = P[k]; F = face_x*T; d = np.zeros(NQ); d[p['t0']//3-1] += T
    if source == 'retained':
        start = p['delivery'] + 3 + delay_m + retained_sale_delay_m
        for i in range(REAL_Q): d[(start+3*i)//3-1] -= F*price*(1-COMM)/REAL_Q
        if landlord == 'waives':   # landlord share on investor units removed from excess settlement (pro rata of base excess profile)
            base = series(k)['exc']; d += LL*F*price*base/base.sum()
    else:                           # 'sold': units taken out of the sold programme, proportional across all collection waves
        base = series(k); f = F/p['sold_total']
        d += -f*base['coll']*price + f*base['comm']          # lost collections, saved commission
        if landlord == 'waives': d += LL*F*price*base['exc']/base['exc'].sum()
    return d
# ---------------- Option B ----------------
GRACE_M = 24; PAY_Q = 12
def pay_quarters(k):
    p = P[k]; s = p['t0'] + GRACE_M + 3
    return [s+3*i for i in range(PAY_Q)]
def inv_flow(k, T, payouts):                       # payouts array (NQ)
    a = -np.array(payouts)*0; a = np.zeros(NQ); a[P[k]['t0']//3-1] -= T; return a + payouts
def s1_fixed(k, T, mult, profile=None):
    q = pay_quarters(k); a = np.zeros(NQ)
    w = np.ones(PAY_Q)/PAY_Q if profile is None else np.array(profile)/np.sum(profile)
    for m, wi in zip(q, w): a[m//3-1] += mult*T*wi
    return a
def s2_pct(k, T, mult, pct, scen, cap_q=None, floor_q=0.0):
    """pct of collections (gross collections of the project) from first payout quarter until mult*T reached. floor/cap per quarter in EGP m."""
    a = np.zeros(NQ); target = mult*T; paid = 0.0
    for m in range(pay_quarters(k)[0], 3*NQ+1, 3):
        i = m//3-1; due = pct*scen['coll'][i]
        if cap_q is not None: due = min(due, cap_q)
        due = max(due, floor_q)
        due = min(due, target-paid)
        a[i] = due; paid += due
        if paid >= target-1e-9: break
    return a
def calibrate_pct(k, T, mult, scen, cap_q, floor_q, lo=0.0, hi=1.0):
    """pct such that the target multiple is reached exactly at the end of the 12-quarter window (base case)."""
    last = pay_quarters(k)[-1]
    def g(pct):
        a = s2_pct(k, T, mult, pct, scen, cap_q, floor_q)
        return a[:last//3].sum() - mult*T
    if g(hi) < 0: return float('nan')
    return brentq(g, lo+1e-9, hi)
def coverage(k, payouts, scen):
    q = pay_quarters(k); out = []
    for m in q:
        i = m//3-1
        out.append((m, scen['net'][i], payouts[i], (scen['net'][i]/payouts[i]) if payouts[i] > 1e-9 else float('inf')))
    return out
def jalour_B(k, T, payouts):
    d = -np.array(payouts); d[P[k]['t0']//3-1] += T; return d

# ---------------- Option B: additional machinery ----------------
def calibrate_min_pct(k, T, mult, scen, cap_q, floor_q):
    """smallest pct whose payouts reach mult*T inside the 12-quarter window (base case)."""
    last = pay_quarters(k)[-1]
    def reach(p): return s2_pct(k, T, mult, p, scen, cap_q, floor_q)[:last//3].sum() >= mult*T-1e-6
    lo, hi = 0.0, 1.0
    if not reach(hi): return float('nan')
    for _ in range(60):
        mid = (lo+hi)/2
        if reach(mid): hi = mid
        else: lo = mid
    return hi
def s3_hybrid(k, T, fixed_mult, total_cap_mult, part_pct, scen):
    """fixed minimum (fixed_mult*T over 12 equal instalments) + part_pct of collections from first payout quarter until total_cap_mult*T paid."""
    a = s1_fixed(k, T, fixed_mult); paid = a.sum(); cap = total_cap_mult*T
    for m in pay_quarters(k):
        i = m//3-1; x = min(part_pct*scen['coll'][i], cap-paid)
        a[i] += x; paid += x
    return a
def calibrate_hybrid(k, T, fixed_mult, target_mult, total_cap_mult, scen):
    last = pay_quarters(k)[-1]; 
    def tot(p): return s3_hybrid(k, T, fixed_mult, total_cap_mult, p, scen)[:last//3].sum()
    lo, hi = 0.0, 1.0
    if tot(hi) < target_mult*T: return float('nan')
    for _ in range(60):
        mid = (lo+hi)/2
        if tot(mid) >= target_mult*T: hi = mid
        else: lo = mid
    return hi
def retained_proceeds(k, scen_price=1.0, delay_m=0):
    p = P[k]; a = np.zeros(NQ); start = p['delivery']+3+delay_m
    for i in range(REAL_Q): a[(start+3*i)//3-1] += p['retained_val']*scen_price*(1-COMM)/REAL_Q
    return a
def cov_stats(k, pay, net, extra=None):
    idx = [m//3-1 for m in pay_quarters(k)]; tot = net + (extra if extra is not None else 0)
    c = [tot[i]/pay[i] for i in idx if pay[i] > 1e-9]
    return dict(min=min(c) if c else float('nan'), n_below=sum(x < 1.5 for x in c), n=len(c),
                agg=sum(tot[i] for i in idx)/max(1e-9, pay[idx[0]:idx[-1]+1].sum()))
def s4_sculpted(k, T, mult, net, extra=None, gate=1.5, max_q=12):
    """Fixed multiple paid in equal amounts across the quarters inside the 12-quarter window in which (net+extra)/instalment >= gate (iterate to fixed point)."""
    idx = [m//3-1 for m in pay_quarters(k)]; tot = net + (extra if extra is not None else 0)
    ok = list(idx)
    for _ in range(30):
        inst = mult*T/len(ok)
        new = [i for i in idx if tot[i]/inst >= gate]
        if new == ok or not new: break
        ok = new
    a = np.zeros(NQ); inst = mult*T/len(ok)
    for i in ok: a[i] = inst
    return a
