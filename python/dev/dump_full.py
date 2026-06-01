import zipfile, sys
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
path = sys.argv[1]
with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find('.//{%s}body' % W)
sect_prs = list(body.iter('{%s}sectPr' % W))
for i, sp in enumerate(sect_prs[:3]):  # show first 3 sectPrs
    pPr = sp.getparent()
    txt = ''
    if pPr is not None and pPr.tag == ('{%s}pPr' % W):
        p = pPr.getparent()
        if p is not None:
            txt = ''.join(t.text or '' for t in p.iter('{%s}t' % W)).strip()
    elif pPr is not None and pPr.tag == ('{%s}body' % W):
        txt = '[body-level]'
    print("=== sectPr[%d] para='%s' ===" % (i, txt[:60]))
    # strip namespaces for readability
    for child in sp:
        tag = child.tag.split('}')[-1]
        attrs = {k.split('}')[-1]: v for k, v in child.attrib.items()}
        print("  <%s %s>" % (tag, attrs))
    print()
