import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn

def get_section_type(sectPr):
    typ_el = sectPr.find(qn('w:type'))
    return typ_el.get(qn('w:val'), 'nextPage') if typ_el is not None else 'nextPage'

def get_orientation(sectPr):
    pgSz = sectPr.find(qn('w:pgSz'))
    if pgSz is None:
        return 'portrait'
    w = int(pgSz.get(qn('w:w'), 0))
    h = int(pgSz.get(qn('w:h'), 0))
    return 'landscape' if w > h else 'portrait'

# Cek OUTPUT file
path = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 19.docx'
doc = Document(path)

print("=== SECTION BREAK DI FILE HASIL ===\n")
print(f"{'Sec':>3} {'Type':>12} {'Orient':>10} {'Start':>6} {'Fmt':>14}  First content")
print("-" * 80)

paras = list(doc.paragraphs)
from utils import DocProcessor

breaks_idx = [i for i, p in enumerate(paras) if DocProcessor._has_sectPr(p._p)]
bounds = [0] + [b + 1 for b in breaks_idx]

for si, sec in enumerate(doc.sections):
    sectPr = sec._sectPr
    stype  = get_section_type(sectPr)
    orient = get_orientation(sectPr)
    pn     = sectPr.find(qn('w:pgNumType'))
    fmt    = pn.get(qn('w:fmt'), 'decimal') if pn is not None else 'decimal'
    start  = pn.get(qn('w:start'), '-') if pn is not None else '-'

    # first content
    s = bounds[si]
    e = bounds[si + 1] if si + 1 < len(bounds) else len(paras)
    first = ''
    for i in range(s, min(s + 10, e)):
        t = paras[i].text.strip()
        if t:
            first = t[:40]
            break

    print(f"{si:>3} {stype:>12} {orient:>10} {start:>6} {fmt:>14}  {first}")

print()
print("=== COLUMN BREAK DI FILE HASIL ===")
col = sum(1 for p in paras for r in p._p.findall('.//' + qn('w:br'))
          if r.get(qn('w:type'), '') == 'column')
print(f"Column break tersisa: {col}")
