"""Remove the third project from the source model using LibreOffice (references update automatically)."""
import sys, os, time, subprocess, shutil, glob
sys.path.insert(0, glob.glob('/root/.claude/skills/synced/*/xlsx/scripts')[0])
from office.soffice import get_soffice_env
import uno
from com.sun.star.beans import PropertyValue
SRC, OUT = sys.argv[1], sys.argv[2]
DEL_ROWS = {'Summary': [5, 12], 'Net Cash Flow -With DP': [6, 13, 20, 27, 34, 41, 48, 55, 62, 69, 101], 'Net Cash Flow -Without DP': [6, 13, 20, 27, 34, 41, 48, 55, 62, 69, 101], 'Cash In Flow Summary': [6, 12, 19, 25, 35, 44, 51], 'Cash Out Detail': [30, 36, 42, 48]}
DEL_COLS = {'Offers': [5], 'Offers DP': [5]}
DEL_SHEETS = ['AT EAST', 'At East Summary']
CLEAR = {'Offers DP': ['I3', 'I4'], 'Cash Out Detail': ['D53', 'E53']}
prof = '/tmp/lo_edit_profile'; shutil.rmtree(prof, ignore_errors=True)
p = subprocess.Popen(['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--invisible', '--norestore', '--accept=socket,host=localhost,port=2002;urp;'], env=get_soffice_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for _ in range(60):
    try:
        local = uno.getComponentContext(); resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
assert ctx, 'no soffice'
smgr = ctx.ServiceManager; desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
def pv(n, v): x = PropertyValue(); x.Name = n; x.Value = v; return x
doc = desktop.loadComponentFromURL('file://' + SRC, '_blank', 0, (pv('Hidden', True),))
sheets = doc.Sheets
for sh, cells in CLEAR.items():
    s = sheets.getByName(sh)
    for a in cells: s.getCellRangeByName(a).clearContents(1 | 2 | 4 | 16)   # value, datetime, string, formula
for sh, rows in DEL_ROWS.items():
    s = sheets.getByName(sh)
    for r in sorted(rows, reverse=True): s.Rows.removeByIndex(r-1, 1)
for sh, cols in DEL_COLS.items():
    s = sheets.getByName(sh)
    for c in sorted(cols, reverse=True): s.Columns.removeByIndex(c-1, 1)
for n in DEL_SHEETS: sheets.removeByName(n)
doc.calculateAll()
doc.storeToURL('file://' + OUT, (pv('FilterName', 'Calc MS Excel 2007 XML'),))
doc.close(True)
try: desktop.terminate()
except Exception: pass
p.wait(timeout=60); print('saved', OUT)
