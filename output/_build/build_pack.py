import sys, os, json, shutil, subprocess; sys.path.insert(0, '.')
import packdata as PD, gen_annex, gen_memo, gen_dd, export_pack
OUT = '/home/user/Jalour_inv_pitch/output'
RECALC = '/root/.claude/skills/synced/8e697198-15eb-43a6-b507-8c3757e6e339_12c2af44-09f8-4a5d-8395-d9d914cc879e/xlsx/scripts/recalc.py'
def build(k, T):
    folder = f'{OUT}/{k}_{T:03d}M'; os.makedirs(folder, exist_ok=True); work = f'/tmp/w/work_{k}_{T}'; os.makedirs(work, exist_ok=True)
    pj = f'{work}/pack.json'; export_pack.export(k, T, pj)
    info = {}
    # annex
    ax = f'{folder}/Financial_Annex.xlsx'; mp = gen_annex.make_annex(k, T, ax); json.dump({'M': mp['M']}, open(f'{work}/annex_map.json', 'w'))
    r = subprocess.run(['python3', RECALC, ax, '180'], capture_output=True, text=True); info['recalc'] = json.loads(r.stdout)
    # deck
    r = subprocess.run(['node', 'build_deck.js', pj, 'market.json', f'{folder}/Deck.pptx'], capture_output=True, text=True); info['deck'] = (r.stdout + r.stderr).strip()
    info['memo'] = gen_memo.build_memo(k, T, f'{folder}/Investment_Memo.docx', work); info['dd'] = gen_dd.build_dd(k, T, f'{folder}/DD_Memo.docx')
    return info
if __name__ == '__main__':
    k, T = sys.argv[1], int(sys.argv[2]); i = build(k, T); print(json.dumps({a: (b if not isinstance(b, dict) else {c: d for c, d in b.items() if c != 'pdf'}) for a, b in i.items()}, default=str)[:600])
