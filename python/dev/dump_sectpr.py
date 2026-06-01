import zipfile, sys
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

path = sys.argv[1]
with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")

root = etree.fromstring(xml)
body = root.find('.//{%s}body' % W)

sect_prs = list(body.iter('{%s}sectPr' % W))
print("Total sectPr: %d\n" % len(sect_prs))

for real_idx, sp in enumerate(sect_prs):
    pPr = sp.getparent()
    txt = ''
    if pPr is not None and pPr.tag == ('{%s}pPr' % W):
        p = pPr.getparent()
        if p is not None:
            txt = ''.join(t.text or '' for t in p.iter('{%s}t' % W)).strip()
    elif pPr is not None and pPr.tag == ('{%s}body' % W):
        txt = '[body-level]'

    type_el = sp.find('{%s}type' % W)
    stype = type_el.get('{%s}val' % W) if type_el is not None else 'nextPage(default)'

    pgSz = sp.find('{%s}pgSz' % W)
    pgMar = sp.find('{%s}pgMar' % W)
    pgNum = sp.find('{%s}pgNumType' % W)
    refs = [e.tag.split('}')[-1] + '/' + e.get('{%s}type' % W, '?')
            for e in sp if 'Reference' in e.tag.split('}')[-1]]

    print("[%02d] type=%-22s para='%s'" % (real_idx, stype, txt[:60]))
    print("     refs=%s" % refs)
    if pgNum is not None:
        attrs = {k.split('}')[-1]: v for k, v in pgNum.attrib.items()}
        print("     pgNumType: %s" % attrs)
    if pgSz is not None:
        attrs = {k.split('}')[-1]: v for k, v in pgSz.attrib.items()}
        print("     pgSz:      %s" % attrs)
    if pgMar is not None:
        attrs = {k.split('}')[-1]: v for k, v in pgMar.attrib.items()}
        print("     pgMar:     %s" % attrs)
    print()
