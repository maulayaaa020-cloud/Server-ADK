"""Trace advance_roman_start untuk Docx 14 — cek apakah sectPr dihapus dari [43]"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from utils import DocProcessor, is_roman_start
from docx.oxml.ns import qn
from lxml import etree

W    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'Docx 14.docx')

doc  = Document(SRC)
body = list(doc.element.body)
proc = DocProcessor(doc, 'Times New Roman', 12)
rsp, _ = proc.scan_zones()
rsp_i = body.index(rsp) if rsp in body else -1
print(f"roman_start_p index: {rsp_i}, text: {repr(DocProcessor._p_text(rsp)[:50])}")

# Cek body[43] sebelum advance
el43 = body[43]
pPr43 = el43.find('{%s}pPr'%W)
sp43 = pPr43.find('{%s}sectPr'%W) if pPr43 is not None else None
print(f"\nSebelum advance: body[43] has sectPr: {sp43 is not None}")

# Run advance_roman_start
new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
new_rsp_i = list(doc.element.body).index(new_rsp) if new_rsp in list(doc.element.body) else -1
print(f"\nSetelah advance: new_rsp_i={new_rsp_i}, use_exact={use_exact}")
print(f"  new_rsp text: {repr(DocProcessor._p_text(new_rsp)[:50])}")

# Cek body[43] setelah advance
body2 = list(doc.element.body)
el43b = body2[43]
pPr43b = el43b.find('{%s}pPr'%W)
sp43b = pPr43b.find('{%s}sectPr'%W) if pPr43b is not None else None
print(f"\nSetelah advance: body[43] has sectPr: {sp43b is not None}")

# Cek body[34] setelah advance
el34b = body2[34]
pPr34b = el34b.find('{%s}pPr'%W)
sp34b = pPr34b.find('{%s}sectPr'%W) if pPr34b is not None else None
print(f"Setelah advance: body[34] '{DocProcessor._p_text(el34b)[:20]}' has sectPr: {sp34b is not None}")

# Simulasi _is_empty_sectPr untuk [43]
def _el_tag(el):
    return el.tag.split('}')[-1] if '}' in el.tag else el.tag
def _txt(el):
    return ''.join(t.text or '' for t in el.iter('{%s}t'%W)).strip()
def _is_empty_sectPr(el):
    if _el_tag(el) != 'p': return False
    if _txt(el): return False
    pPr = el.find('{%s}pPr'%W)
    return pPr is not None and pPr.find('{%s}sectPr'%W) is not None

print(f"\n_is_empty_sectPr(body[43]): {_is_empty_sectPr(body[43])}")
print(f"_is_empty_sectPr(body[44]): {_is_empty_sectPr(body[44])}")
