"""Docx 6 full structure investigation"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii')

def _has_pgbr(el):
    return any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
def _has_sectPr(el):
    pPr = el.find("{%s}pPr" % W)
    return pPr is not None and pPr.find("{%s}sectPr" % W) is not None
def txt(el, n=60):
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:n]
def tag(el):
    return el.tag.split('}')[-1] if '}' in el.tag else el.tag

doc  = Document(os.path.join(SRC, 'Docx 6.docx'))
proc = DocProcessor(doc, 'Times New Roman', 12)
rsp, bab = proc.scan_zones()
body = list(doc.element.body)
rsp_idx = body.index(rsp)
bab_idx = [body.index(p) for p in bab]

# Find all page breaks in input
print("All page breaks in INPUT Docx 6:")
for i, el in enumerate(body):
    if _has_pgbr(el) or _has_sectPr(el):
        spr = '[sectPr]' if _has_sectPr(el) else '[pgBr]  '
        mk = ' <--RSP' if i == rsp_idx else ''
        mk += ' <--BAB' if i in bab_idx else ''
        print("  [%3d] %s %r%s" % (i, spr, txt(el), mk))

print()
print("Structure [16..60]:")
for i in range(16, min(61, len(body))):
    el = body[i]
    t  = tag(el)
    mk = ' <--RSP' if i == rsp_idx else ''
    mk += ' <--BAB' if i in bab_idx else ''
    spr = '[sectPr]' if _has_sectPr(el) else ('[pgBr]  ' if _has_pgbr(el) else '        ')
    ct  = 'C' if proc._p_has_content(el) else ' '
    print("  [%3d] %s%s %s %r%s" % (i, ct, t, spr, txt(el), mk))
