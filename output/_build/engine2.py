"""Revised structuring engine (v2): Option A from sold launch tranches at launch list price; Option B 2.3x; nothing ring-fenced."""
import json, numpy as np
from engine import *          # series(), npv(), irr(), P, MONTHS, NQ, T_Y, pad, etc.
TR = json.load(open(__file__.replace('engine2.py', 'tranches.json')))
M_B2 = 2.3; FACE_X = 2.0
def padv(a): return pad(a)
def launch(k, part='all'):
    """first sales tranches that carry the launch price list; returns dict with flows, values, final/launch ratios."""
    tr = TR[k]['tranches']; p0 = (tr[0]['price_comm'], tr[0]['price_admin'])
    L = [t for t in tr if (t['price_comm'], t['price_admin']) == p0]
    fin = (tr[-1]['price_comm'], tr[-1]['price_admin'])
    Vc = sum(t['sales_comm'] for t in L)/1000; Va = sum(t['sales_admin'] for t in L)/1000
    flow = {'all': sum(np.array(t['flow']) for t in L), 'comm': sum(np.array(t['flow_comm']) for t in L), 'admin': sum(np.array(t['flow_admin']) for t in L)}
    return dict(tr=L, p0=p0, fin=fin, Vc=Vc, Va=Va, flow={a: padv(b) for a, b in flow.items()},
                months=[t['month'] for t in L], r_c=fin[0]/p0[0], r_a=fin[1]/p0[1])
def mix(k, part):
    Lh = launch(k); V = {'all': Lh['Vc']+Lh['Va'], 'comm': Lh['Vc'], 'admin': Lh['Va']}[part]
    wc = {'all': Lh['Vc']/(Lh['Vc']+Lh['Va']), 'comm': 1.0, 'admin': 0.0}[part]
    return Lh, V, wc
def appreciation(k, part='all', g=None):
    Lh, V, wc = mix(k, part)
    if g is not None: return 1+g
    return wc*Lh['r_c'] + (1-wc)*Lh['r_a']
def optA_inv(k, T, face_x=FACE_X, price=1.0, delay_m=0, friction=FRICTION, part='all', g=None):
    p = P[k]; F = face_x*T; a = np.zeros(NQ); a[p['t0']//3-1] -= T
    start = p['delivery']+3+delay_m; val = F*appreciation(k, part, g)*price*(1-friction)
    for i in range(REAL_Q): a[(start+3*i)//3-1] += val/REAL_Q
    return a
def optA_jal(k, T, variant='launch', face_x=FACE_X, part='all'):
    """incremental cash flow to Jalour. Jalour bears the landlord 35% (landlord payments unchanged, as in the model)."""
    p = P[k]; F = face_x*T; d = np.zeros(NQ); d[p['t0']//3-1] += T; b = series(k)
    if variant == 'launch':
        Lh, V, wc = mix(k, part); f = F/V
        d -= f*Lh['flow'][part]
        # commission (4% of sales value) saved, paid 12 months after each sale quarter
        for t in Lh['tr']:
            sv = {'all': t['sales_comm']+t['sales_admin'], 'comm': t['sales_comm'], 'admin': t['sales_admin']}[part]/1000
            d[(t['month']+12)//3-1] += f*sv*COMM
        return d, f
    if variant == 'prorata':       # same area (face at launch list) taken pro rata from every tranche, i.e. at average realised prices
        Lh = launch(k); tot_c = sum(t['sales_comm'] for t in TR[k]['tranches'])/1000; tot_a = sum(t['sales_admin'] for t in TR[k]['tranches'])/1000
        wc = Lh['Vc']/(Lh['Vc']+Lh['Va']); avg_ratio = wc*(tot_c/(tot_c+tot_a))*0  # placeholder replaced below
        area_c = sum(t['area_comm'] for t in TR[k]['tranches']); area_a = sum(t['area_admin'] for t in TR[k]['tranches'])
        avg_c = tot_c*1000/area_c; avg_a = tot_a*1000/area_a
        Feff = F*(wc*avg_c/Lh['p0'][0] + (1-wc)*avg_a/Lh['p0'][1]); f = Feff/(tot_c+tot_a)
        d += -f*b['coll'] + f*b['comm']
        return d, f
