import sys, os, json, subprocess; sys.path.insert(0, '.')
import packdata as PD, gen_annex_combined, gen_memo_combined, gen_dd_combined, gen_nda, export_pack
OUT = '/home/user/Jalour_inv_pitch/output'; FOLDER = f'{OUT}/COMBINED_125M'
RECALC = '/root/.claude/skills/synced/8e697198-15eb-43a6-b507-8c3757e6e339_12c2af44-09f8-4a5d-8395-d9d914cc879e/xlsx/scripts/recalc.py'
def build(Te=62.5):
    os.makedirs(FOLDER, exist_ok=True); work = '/tmp/w/work_comb'; os.makedirs(work, exist_ok=True); info = {}
    ax = f'{FOLDER}/Financial_Annex.xlsx'; r = gen_annex_combined.make(Te, ax, work); json.dump({'M': r['M']}, open(f'{work}/annex_map.json', 'w'))
    p = subprocess.run(['python3', RECALC, ax, '300'], capture_output=True, text=True); info['recalc'] = p.stdout[:300]
    C = PD.build_combined(Te); cj = f'{work}/comb.json'; json.dump(C, open(cj, 'w'), default=export_pack.conv)
    p = subprocess.run(['node', 'build_deck_combined.js', cj, 'market.json', f'{FOLDER}/Deck.pptx'], capture_output=True, text=True); info['deck'] = (p.stdout + p.stderr).strip()
    info['memo'] = gen_memo_combined.build_memo_combined(f'{FOLDER}/Investment_Memo.docx', work, Te)['pages']
    info['dd'] = gen_dd_combined.build_dd_combined(f'{FOLDER}/DD_Memo.docx', Te)['pages']
    gen_nda.build(f'{FOLDER}/NDA_Template.docx', ["Green Square", "L'avenir"])
    for f in ('Investment_Memo', 'DD_Memo', 'NDA_Template'): subprocess.run(['python3', 'fix_docx.py', f'{FOLDER}/{f}.docx'])
    for f in ('Deck.pptx', 'Investment_Memo.docx', 'DD_Memo.docx', 'NDA_Template.docx'): subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', FOLDER, f'{FOLDER}/{f}'], capture_output=True)
    return info
if __name__ == '__main__': print(build())
