import sys; sys.path.insert(0, '.')
import build_pack, qa, subprocess
for k in ('GS', 'LA'):
    i = build_pack.build(k, 125)
    subprocess.run(['python3', 'fix_docx.py', f'../{k}_125M/Investment_Memo.docx', f'../{k}_125M/DD_Memo.docx'])
    print(k, 125, i['recalc']['total_errors'], i['deck'], i['memo']['pages'], i['dd']['pages'], 'qa', qa.run(k, 125), flush=True)
for k in ('GS', 'LA'):
    for T in (50, 75, 100): print(k, T, 'qa', qa.run(k, T), flush=True)
