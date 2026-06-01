"""Trace full flow untuk Docx 14 — advance + insert_breaks"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from utils import DocProcessor, is_roman_start
from docx.oxml.ns import qn

W    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'Docx 14.docx')

doc  = Document(SRC)
proc = DocProcessor(doc, 'Times New Roman', 12)
rsp, bab_p_list = proc.scan_zones()

body = list(doc.element.body)
rsp_i = body.index(rsp) if rsp in body else -1
print(f"roman_start_p index: {rsp_i}")

# Run advance_roman_start
new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
new_body = list(doc.element.body)
new_rsp_i = new_body.index(new_rsp) if new_rsp in new_body else -1
print(f"after advance: new_rsp_i={new_rsp_i}, use_exact={use_exact}")

# Check sectPr positions after advance
def _has_sectPr(el):
    pPr = el.find('{%s}pPr'%W)
    return pPr is not None and pPr.find('{%s}sectPr'%W) is not None

sec_after_advance = [(i, DocProcessor._p_text(e)[:30]) for i,e in enumerate(new_body[:50]) if e.tag.endswith('}p') and _has_sectPr(e)]
print(f"sectPr positions [0..49] after advance: {sec_after_advance}")

# Run insert_breaks
if new_rsp is not rsp:
    roman_start_p = new_rsp
else:
    roman_start_p = rsp

returned_rsp = proc.insert_breaks(roman_start_p, bab_p_list, exact_roman_start=use_exact)
print(f"\nafter insert_breaks: returned_rsp text: {repr(DocProcessor._p_text(returned_rsp)[:30])}")

# Check sectPr positions after insert_breaks
final_body = list(doc.element.body)
sec_after_ib = [(i, repr(DocProcessor._p_text(e)[:30])) for i,e in enumerate(final_body[:50]) if e.tag.endswith('}p') and _has_sectPr(e)]
print(f"sectPr positions [0..49] after insert_breaks: {sec_after_ib}")
