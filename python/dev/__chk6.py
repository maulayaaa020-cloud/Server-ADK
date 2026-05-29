"""Check Docx 6: input structure + output section detail"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii')
OUT  = os.path.join(SRC, 'hasil', 'Docx 6_v2.docx')

def _has_pgbr(el):
    return any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
def _has_sectPr(el):
    pPr = el.find("{%s}pPr" % W)
    return pPr is not None and pPr.find("{%s}sectPr" % W) is not None
def txt(el, n=45):
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:n]
def tag(el):
    return el.tag.split('}')[-1] if '}' in el.tag else el.tag

# --- INPUT ---
doc  = Document(os.path.join(SRC, 'Docx 6.docx'))
proc = DocProcessor(doc, 'Times New Roman', 12)
rsp, bab = proc.scan_zones()
body = list(doc.element.body)
rsp_idx = body.index(rsp)
bab_idx = [body.index(p) for p in bab]
bb = sum(1 for el in body[:rsp_idx] if (_has_pgbr(el) or _has_sectPr(el)))

new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 2)
new_idx = body.index(new_rsp)

print("INPUT Docx 6: bb=%d, rsp=[%d]=%r" % (bb, rsp_idx, txt(rsp)))
print("  advance -> new=[%d]=%r  exact=%s" % (new_idx, txt(new_rsp), use_exact))
print("  bab: %s" % str(bab_idx[:3]))
print()

# Show structure [0..new_idx+5]
hi = min(new_idx + 8, len(body))
print("Structure [0..%d]:" % (hi - 1))
for i in range(hi):
    el = body[i]
    t  = tag(el)
    mk = ''
    if i == rsp_idx: mk += ' <--RSP'
    if i == new_idx: mk += ' <--NEW'
    if i in bab_idx: mk += ' <--BAB'
    spr = '[sectPr]' if _has_sectPr(el) else ('[pgBr]  ' if _has_pgbr(el) else '        ')
    print("  [%3d] %-4s %s %r%s" % (i, t, spr, txt(el), mk))

# --- OUTPUT ---
print()
print("OUTPUT sections:")
if os.path.exists(OUT):
    doc2 = Document(OUT)
    for i, sec in enumerate(doc2.sections):
        pn    = sec._sectPr.find(qn('w:pgNumType'))
        fmt   = pn.get(qn('w:fmt'), 'decimal') if pn is not None else 'decimal'
        start = pn.get(qn('w:start'), None) if pn is not None else None
        print("  sec[%d]: fmt=%s  start=%s" % (i, fmt, start))

    print()
    print("OUTPUT: paragraphs with sectPr (first 5):")
    body2 = list(doc2.element.body)
    count = 0
    for i, el in enumerate(body2):
        if _has_sectPr(el):
            print("  [%3d] %r" % (i, txt(el)))
            count += 1
            if count >= 5: break
else:
    print("  (file not found)")
