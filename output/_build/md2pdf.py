import sys, glob, os, subprocess
CH = '/opt/pw-browsers/chromium-1194/chrome-linux/chrome'
CSS = '/tmp/md2pdf_css.html'
open(CSS, 'w').write("<style>@page{size:A4 landscape;margin:12mm}body{font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#1a1a1a;max-width:none;padding:0}h1{color:#1B2F4B;font-size:18pt}h2{color:#1B2F4B;font-size:13pt;margin-top:16px}table{border-collapse:collapse;width:100%;margin:8px 0;font-size:8.5pt}th{background:#1B2F4B;color:#fff;text-align:left}th,td{border:1px solid #C9CFD8;padding:3px 5px;vertical-align:top}tr:nth-child(even) td{background:#F3F5F8}</style>")
def conv(md, remove=True):
    b = md[:-3]; subprocess.run(['pandoc', md, '-f', 'gfm', '-t', 'html', '-s', '--metadata', 'title=' + os.path.basename(b).replace('_', ' '), '-H', CSS, '-o', b + '.html'], check=True)
    subprocess.run([CH, '--headless', '--no-sandbox', '--disable-gpu', '--no-pdf-header-footer', f'--print-to-pdf={b}.pdf', 'file://' + os.path.abspath(b + '.html')], capture_output=True)
    os.remove(b + '.html')
    if remove: os.remove(md)
if __name__ == '__main__':
    for f in sys.argv[1:]: conv(f)
