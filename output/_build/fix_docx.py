import zipfile, shutil, re, sys, glob, os
def fix(p):
    tmp = p + '.tmp'; zin = zipfile.ZipFile(p); zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename == 'word/settings.xml':
            s = data.decode('utf8'); s = re.sub(r'<w:zoom(?![^>]*w:percent)([^>]*)/>', r'<w:zoom w:percent="100"\1/>', s); data = s.encode('utf8')
        if it.filename == 'word/document.xml':
            from lxml import etree
            W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'; root = etree.fromstring(data)
            order = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
            for pr in root.iter('{%s}tblPr' % W):
                kids = list(pr); kids.sort(key=lambda e: order.index(etree.QName(e).localname) if etree.QName(e).localname in order else 99)
                for e in kids: pr.remove(e)
                for e in kids: pr.append(e)
            data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(it, data)
    zout.close(); zin.close(); os.replace(tmp, p)
for p in sys.argv[1:]: fix(p)
