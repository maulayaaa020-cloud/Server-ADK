import zipfile, sys
from lxml import etree

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WPD = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def dump(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find('.//{%s}body' % W)
    print('=== %s ===' % label)
    for i, el in enumerate(body):
        tag = el.tag.split('}')[-1]
        if tag != 'p':
            continue
        txt = ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()
        has_sect = el.find('.//{%s}sectPr' % W) is not None
        has_pgbr = any(br.get('{%s}type' % W) == 'page' for br in el.iter('{%s}br' % W))
        has_img  = (el.find('.//{%s}inline' % WPD) is not None or
                    el.find('.//{%s}anchor' % WPD) is not None)
        marks = ''
        if has_sect: marks += '[SECT]'
        if has_pgbr: marks += '[PGBR]'
        if has_img:  marks += '[IMG]'
        short = txt[:70] + ('...' if len(txt) > 70 else '')
        print('[%03d] %-14s %s' % (i, marks, short))

dump(sys.argv[1], 'ORIGINAL')
print()
dump(sys.argv[2], 'HASIL')
