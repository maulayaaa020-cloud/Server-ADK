"""Debug Docx 1 cover 1 dimulai dari ii"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii')
OUT  = os.path.join(SRC, 'hasil', 'Docx 1_c1.docx')

def _has_pgbr(el):
    return any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
def _has_sectPr(el):
    pPr = el.find("{%s}pPr" % W)
    return pPr is not None and pPr.find("{%s}sectPr" % W) is not None
def txt(el, n=50):
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:n]
def tag(el):
    return el.tag.split('}')[-1] if '}' in el.tag else el.tag

doc  = Document(os.path.join(SRC, 'Docx 1.docx'))
proc = DocProcessor(doc, 'Times New Roman', 12)
rsp, bab = proc.scan_zones()
body = list(doc.element.body)
rsp_idx = body.index(rsp)
bab_idx = [body.index(p) for p in bab]
bb = sum(1 for el in body[:rsp_idx] if (_has_pgbr(el) or _has_sectPr(el)))

new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
new_idx = body.index(new_rsp)

print("INPUT Docx 1 (cover 1): bb=%d, rsp=[%d]=%r" % (bb, rsp_idx, txt(rsp)))
print("  advance(num_cover=1) -> new=[%d]=%r  exact=%s" % (new_idx, txt(new_rsp), use_exact))
print("  bab: %s" % str(bab_idx[:4]))
print()

# Semua breaks di input
print("All breaks in INPUT:")
for i, el in enumerate(body):
    if _has_pgbr(el) or _has_sectPr(el):
        spr = '[sectPr]' if _has_sectPr(el) else '[pgBr]  '
        mk = ' <--RSP' if i == rsp_idx else ''
        mk += ' <--NEW' if i == new_idx else ''
        mk += ' <--BAB' if i in bab_idx else ''
        print("  [%3d] %s %r%s" % (i, spr, txt(el), mk))

print()
lo = max(0, rsp_idx - 3)
hi = min(len(body), new_idx + 5)
print("Structure [%d..%d]:" % (lo, hi-1))
for i in range(lo, hi):
    el = body[i]
    t  = tag(el)
    mk = ' <--RSP' if i == rsp_idx else ''
    mk += ' <--NEW' if i == new_idx else ''
    mk += ' <--BAB' if i in bab_idx else ''
    spr = '[sectPr]' if _has_sectPr(el) else ('[pgBr]  ' if _has_pgbr(el) else '        ')
    print("  [%3d] %-4s %s %r%s" % (i, t, spr, txt(el), mk))

# Output sections
print()
print("OUTPUT sections:")
if os.path.exists(OUT):
    doc2 = Document(OUT)
    for i, sec in enumerate(doc2.sections):
        pn    = sec._sectPr.find(qn('w:pgNumType'))
        fmt   = pn.get(qn('w:fmt'), 'decimal') if pn is not None else 'decimal'
        start = pn.get(qn('w:start'), None) if pn is not None else None
        print("  sec[%d]: fmt=%s  start=%s" % (i, fmt, start))
