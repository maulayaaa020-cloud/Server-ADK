import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

path = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
doc = Document(path)
paras = list(doc.paragraphs)

print("=== JENIS BREAK DI DOKUMEN ASLI ===\n")

sect_continuous = []
sect_nextpage   = []
sect_evenpage   = []
sect_oddpage    = []
sect_other      = []
col_breaks      = []
page_breaks     = []

for i, p in enumerate(paras):
    pPr = p._p.find(qn('w:pPr'))
    # Section breaks
    if pPr is not None:
        sectPr = pPr.find(qn('w:sectPr'))
        if sectPr is not None:
            pgSz = sectPr.find(qn('w:pgSz'))
            typ_el = sectPr.find(qn('w:type'))
            typ = typ_el.get(qn('w:val'), 'nextPage') if typ_el is not None else 'nextPage'
            ctx = repr(p.text.strip()[:40]) if p.text.strip() else '(empty)'
            entry = (i, typ, ctx)
            if typ == 'continuous':
                sect_continuous.append(entry)
            elif typ == 'nextPage':
                sect_nextpage.append(entry)
            elif typ == 'evenPage':
                sect_evenpage.append(entry)
            elif typ == 'oddPage':
                sect_oddpage.append(entry)
            else:
                sect_other.append(entry)

    # Column breaks & page breaks (inside runs)
    for r in p._p.findall('.//' + qn('w:br')):
        btype = r.get(qn('w:type'), 'textWrapping')
        ctx = repr(p.text.strip()[:40]) if p.text.strip() else '(empty)'
        if btype == 'column':
            col_breaks.append((i, ctx))
        elif btype == 'page':
            page_breaks.append((i, ctx))

print(f"Section Break (Continuous) : {len(sect_continuous)}")
print(f"Section Break (Next Page)  : {len(sect_nextpage)}")
print(f"Section Break (Even Page)  : {len(sect_evenpage)}")
print(f"Section Break (Odd Page)   : {len(sect_oddpage)}")
print(f"Section Break (Other)      : {len(sect_other)}")
print(f"Column Break               : {len(col_breaks)}")
print(f"Page Break (manual)        : {len(page_breaks)}")
print()

print("--- Section Break (Continuous) ---")
for i, typ, ctx in sect_continuous:
    print(f"  para[{i:3d}]: {ctx}")

print()
print("--- Section Break (Next Page) ---")
for i, typ, ctx in sect_nextpage:
    print(f"  para[{i:3d}]: {ctx}")

print()
print("--- Column Breaks ---")
for i, ctx in col_breaks:
    print(f"  para[{i:3d}]: {ctx}")
