import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

path = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
doc  = Document(path)
paras = list(doc.paragraphs)

print("=== SEMUA SECTION BREAK DI DOKUMEN BARU ===\n")
for i, p in enumerate(paras):
    if not DocProcessor._has_sectPr(p._p):
        continue
    pPr   = p._p.find(qn('w:pPr'))
    sectPr = pPr.find(qn('w:sectPr'))
    pgSz  = sectPr.find(qn('w:pgSz'))
    typ_el = sectPr.find(qn('w:type'))
    typ   = typ_el.get(qn('w:val'), 'nextPage') if typ_el is not None else 'nextPage'
    has_content = bool(p.text.strip())
    w = pgSz.get(qn('w:w'), '?') if pgSz is not None else '?'
    h = pgSz.get(qn('w:h'), '?') if pgSz is not None else '?'
    print(f"para[{i:3d}] type={typ:12s}  pgSz={w}x{h}  content={has_content}  text={repr(p.text.strip()[:40])}")

print()
print("=== BODY-LEVEL sectPr ===")
body_sectPr = doc.element.body.find(qn('w:sectPr'))
if body_sectPr is not None:
    pgSz = body_sectPr.find(qn('w:pgSz'))
    w = pgSz.get(qn('w:w'), '?') if pgSz is not None else '?'
    h = pgSz.get(qn('w:h'), '?') if pgSz is not None else '?'
    print(f"body sectPr: pgSz={w}x{h}")
