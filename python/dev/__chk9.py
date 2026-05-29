"""Quick check: Docx 9 output section structure"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
OUT  = os.path.join(ROOT, 'test_files', 'hasil', 'Docx 9_p3.docx')

doc = Document(OUT)

def txt(el, n=40):
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:n]

print("Sections in Docx 9 output:")
for i, sec in enumerate(doc.sections):
    pn    = sec._sectPr.find(qn('w:pgNumType'))
    fmt   = pn.get(qn('w:fmt'), 'decimal') if pn is not None else 'decimal'
    start = pn.get(qn('w:start'), None) if pn is not None else None
    print(f"  sec[{i}]: fmt={fmt}  start={start}")

print()
print("Paragraphs with sectPr (section boundaries):")
body = list(doc.element.body)
for i, el in enumerate(body):
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    if tag == 'p':
        pPr = el.find("{%s}pPr" % W)
        if pPr is not None and pPr.find("{%s}sectPr" % W) is not None:
            print(f"  [{i:3d}] {txt(el)!r}")
    if any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W)):
        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if tag == 'p':
            print(f"  [{i:3d}] [pgBr] {txt(el)!r}")
