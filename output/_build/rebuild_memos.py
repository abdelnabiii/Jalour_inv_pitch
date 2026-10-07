import sys, subprocess; sys.path.insert(0, '.')
import gen_memo, gen_dd, qa
for k in ('GS', 'LA'):
    for T in (50, 75, 100, 125):
        gen_memo.build_memo(k, T, f'../{k}_{T:03d}M/Investment_Memo.docx', f'/tmp/w/work_{k}_{T}')
        gen_dd.build_dd(k, T, f'../{k}_{T:03d}M/DD_Memo.docx')
        subprocess.run(['python3', 'fix_docx.py', f'../{k}_{T:03d}M/Investment_Memo.docx', f'../{k}_{T:03d}M/DD_Memo.docx'])
        print(k, T, qa.run(k, T), flush=True)
