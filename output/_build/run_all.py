import sys; sys.path.insert(0, '.')
import build_pack, qa, json
res = {}
for k in ('GS', 'LA'):
    for T in (50, 75, 100):
        if len(sys.argv) > 1 and f'{k}_{T}' not in sys.argv[1:]: continue
        i = build_pack.build(k, T); f = qa.run(k, T); res[f'{k}_{T}'] = dict(recalc=i['recalc']['total_errors'], formulas=i['recalc'].get('total_formulas'), deck=i['deck'], memo_pages=i['memo']['pages'], dd_pages=i['dd']['pages'], qa_failures=f)
        print(k, T, res[f'{k}_{T}'], flush=True)
json.dump(res, open('/tmp/w/run_all.json', 'w'))
