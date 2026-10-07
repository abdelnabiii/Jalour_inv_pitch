import sys, subprocess, glob, os, shutil
from PIL import Image
SK = glob.glob('/root/.claude/skills/synced/*/pptx/scripts/office/soffice.py')[0]
def render(path, outdir, dpi=70, per=6, cols=2):
    os.makedirs(outdir, exist_ok=True)
    for f in glob.glob(outdir + '/p-*.jpg'): os.remove(f)
    subprocess.run(['python3', SK, '--headless', '--convert-to', 'pdf', '--outdir', outdir, path], capture_output=True)
    pdf = os.path.join(outdir, os.path.basename(path).rsplit('.', 1)[0] + '.pdf')
    subprocess.run(['pdftoppm', '-jpeg', '-r', str(dpi), pdf, outdir + '/p'], capture_output=True)
    fs = sorted(glob.glob(outdir + '/p-*.jpg')); ims = [Image.open(f) for f in fs]; w, h = ims[0].size; out = []
    for i in range(0, len(ims), per):
        sub = ims[i:i+per]; rows = (len(sub)+cols-1)//cols; sh = Image.new('RGB', (w*cols, h*rows), 'white')
        for j, im in enumerate(sub): sh.paste(im, ((j % cols)*w, (j//cols)*h))
        p = f'{outdir}/sheet_{i//per}.jpg'; sh.save(p); out.append(p)
    return out, len(ims)
if __name__ == '__main__':
    print(render(sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 70, int(sys.argv[4]) if len(sys.argv) > 4 else 6, int(sys.argv[5]) if len(sys.argv) > 5 else 2))
