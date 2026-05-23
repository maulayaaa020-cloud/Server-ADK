import sys
from docx import Document
from docx.oxml.ns import qn

def has_sectPr(p_elem):
    pPr = p_elem.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:sectPr')) is not None

def has_page_break(p_elem):
    for br in p_elem.iter(qn('w:br')):
        if br.get(qn('w:type')) == 'page':
            return True
    return False

def has_page_break_before_in_pPr(p_elem):
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is None:
        return False
    pb = pPr.find(qn('w:pageBreakBefore'))
    if pb is None:
        return False
    return pb.get(qn('w:val'), 'true') not in ('false', '0')

path = sys.argv[1]
doc = Document(path)
all_paras = list(doc.paragraphs)

print(f"Total paragraf: {len(all_paras)}")
print(f"Total sections: {len(doc.sections)}")
print()
print("=== SEMUA PARAGRAF (termasuk kosong) ===")
print(f"{'idx':>5} | {'style':<30} | {'sect':^4} | {'pgbr':^4} | {'pgbrB':^5} | text[:70]")
print("-"*120)

start = int(sys.argv[2]) if len(sys.argv) > 2 else 0
end   = int(sys.argv[3]) if len(sys.argv) > 3 else len(all_paras)

for i, p in enumerate(all_paras[start:end], start=start):
    text  = p.text.strip()
    style = p.style.name if p.style else "-"
    sect  = "Y" if has_sectPr(p._p) else ""
    pgbr  = "Y" if has_page_break(p._p) else ""
    pgbrB = "Y" if has_page_break_before_in_pPr(p._p) else ""
    print(f"{i:5d} | {style:<30} | {sect:^4} | {pgbr:^4} | {pgbrB:^5} | {repr(text[:70])}")
