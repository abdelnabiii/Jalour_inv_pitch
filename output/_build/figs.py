import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, numpy as np, os
NAVY, BRASS, TEAL, GREY, RED, MID = '#17324D', '#B8893B', '#2A7F83', '#6B7280', '#A23B3B', '#D5DAE1'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9, 'axes.spines.top': False, 'axes.spines.right': False, 'axes.edgecolor': '#9AA3AF', 'axes.labelcolor': '#374151', 'xtick.color': '#374151', 'ytick.color': '#374151'})
def base(figsize=(7.2, 3.6)):
    f, ax = plt.subplots(figsize=figsize, dpi=200); ax.grid(axis='y', color=MID, lw=0.6); ax.set_axisbelow(True); return f, ax
def make(D, outdir):
    os.makedirs(outdir, exist_ok=True); P = {}
    a = D['annual'][:10]; yrs = [f"Y{r['year']}" for r in a]; x = np.arange(len(a)); w = 0.26
    f, ax = base()
    ax.bar(x-w, [r['coll'] for r in a], w, color=NAVY, label='Collections'); ax.bar(x, [-r['land'] for r in a], w, color=BRASS, label='Landlord payments'); ax.bar(x+w, [-r['cost'] for r in a], w, color=GREY, label='Construction, commission, SG&A')
    ax.plot(x, [r['cum'] for r in a], color=RED, lw=2, marker='o', ms=4, label='Cumulative net cash flow'); ax.axhline(0, color='#9AA3AF', lw=0.8)
    ax.set_xticks(x); ax.set_xticklabels(yrs); ax.set_ylabel('EGP million'); ax.legend(frameon=False, ncol=2, fontsize=8, loc='upper right'); f.tight_layout(); P['cash'] = os.path.join(outdir, 'fig_cash.png'); f.savefig(P['cash']); plt.close(f)
    tr = D['tranches']; f, ax = base((7.2, 3.2)); xs = np.arange(len(tr))
    ax.bar(xs, [t['comm'] for t in tr], 0.6, color=BRASS, label='Retail'); ax.bar(xs, [t['admin'] for t in tr], 0.6, bottom=[t['comm'] for t in tr], color=NAVY, label='Offices')
    ax.set_xticks(xs); ax.set_xticklabels([f"M{t['month']}" for t in tr]); ax.set_ylabel('EGP million'); ax.legend(frameon=False, fontsize=8); f.tight_layout(); P['sales'] = os.path.join(outdir, 'fig_sales.png'); f.savefig(P['sales']); plt.close(f)
    M = np.array(D['series']['months']); k = D['k']; t0 = 3 if k == 'GS' else 12
    A = np.cumsum(D['A']['inv']['Base']['flows']); Ad = np.cumsum(D['A']['inv']['Downside']['flows']); B = np.cumsum(np.array([-(D['T'] if m == t0 else 0) + p for m, p in zip(M, D['B']['S']['S1']['res']['Base']['pay'])]))
    n = 28; f, ax = base((7.2, 3.4))
    ax.plot(M[:n], A[:n], color=NAVY, lw=2, label='Option A base'); ax.plot(M[:n], Ad[:n], color=NAVY, lw=1.6, ls='--', label='Option A downside'); ax.plot(M[:n], B[:n], color=BRASS, lw=2, label='Option B')
    ax.axhline(0, color='#9AA3AF', lw=0.8); ax.set_xlabel('Project month'); ax.set_ylabel('Cumulative investor cash flow, EGP million'); ax.legend(frameon=False, fontsize=8); f.tight_layout(); P['inv'] = os.path.join(outdir, 'fig_inv.png'); f.savefig(P['inv']); plt.close(f)
    cases = [('Base', D['sens']['price'][1.0]), ('Prices -10%', D['sens']['price'][0.9]), ('Prices +10%', D['sens']['price'][1.1]), ('Sales delay 12m', D['sens']['delay'][4]), ('Cost +10%', D['sens']['cost'][1.1]), ('Slip 2 quarters', D['sens']['slip'][2]), ('Delivery pay at handover', D['proj']['Delivery payments at handover'])]
    f, ax = base((7.2, 3.3)); ys = np.arange(len(cases))[::-1]
    ax.barh(ys, [c[1]['npv'] for c in cases], 0.55, color=[NAVY if c[1]['npv'] >= 0 else RED for c in cases])
    for y, c in zip(ys, cases): ax.text(c[1]['npv'] + (6 if c[1]['npv'] >= 0 else -6), y, f"{c[1]['npv']:,.0f}", va='center', ha='left' if c[1]['npv'] >= 0 else 'right', fontsize=8, color='#1B1F2A')
    ax.set_yticks(ys); ax.set_yticklabels([c[0] for c in cases]); ax.set_xlabel('NPV of Jalour net cash flow at 14%, EGP million'); ax.grid(axis='x', color=MID, lw=0.6); ax.grid(axis='y', visible=False); f.tight_layout(); P['sens'] = os.path.join(outdir, 'fig_sens.png'); f.savefig(P['sens']); plt.close(f)
    return P
