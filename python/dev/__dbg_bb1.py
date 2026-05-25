"""Cek bb untuk semua file cover 1 dimulai dari ii"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii')

def _has_break(el):
    if any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W)): return True
    pPr = el.find("{%s}pPr" % W)
    return pPr is not None and pPr.find("{%s}sectPr" % W) is not None

def txt(el, n=35):
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:n]

print("%-12s  %3s  %-35s" % ("File", "bb", "rsp_txt"))
print("-" * 60)
for fname in sorted(f for f in os.listdir(SRC) if f.endswith('.docx')):
    doc  = Document(os.path.join(SRC, fname))
    proc = DocProcessor(doc, 'Times New Roman', 12)
    rsp, _ = proc.scan_zones()
    body   = list(doc.element.body)
    rsp_idx = body.index(rsp)
    bb = sum(1 for el in body[:rsp_idx] if _has_break(el))
    flag = " ***" if bb > 1 else ""
    print("%-12s  %3d  %-35r%s" % (fname, bb, txt(rsp), flag))
