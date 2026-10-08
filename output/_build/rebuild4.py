import sys, subprocess; sys.path.insert(0, '.')
import gen_memo, gen_dd, qa
for k in ('GS', 'LA'):
    for T in (50, 75, 100, 125):
        f = f'../{k}_{T:03d}M'
        gen_memo.build_memo(k, T, f'{f}/Investment_Memo.docx', f'/tmp/w/work_{k}_{T}')
        subprocess.run(['python3', 'fix_docx.py', f'{f}/Investment_Memo.docx'])
        print(k, T, qa.run(k, T), flush=True)
