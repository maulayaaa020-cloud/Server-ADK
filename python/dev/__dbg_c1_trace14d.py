"""Trace insert_break_before_xml detail: apa yang menjadi target _attach_sectPr"""
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
proc = DocProcessor(doc, 'Times New Roman', 12)
rsp, bab_p_list = proc.scan_zones()

new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
new_body = list(doc.element.body)
tgt_idx = new_body.index(new_rsp)
print(f"target_p at [{tgt_idx}]")

# Print XML of body[42..44]
for i in range(42, min(46, len(new_body))):
    el = new_body[i]
    xml_str = etree.tostring(el, pretty_print=True).decode()
    txt = DocProcessor._p_text(el)
    hc = DocProcessor._p_has_content(el)
    hs = DocProcessor._has_sectPr(el)
    print(f"\n[{i}] text={repr(txt[:30])}, has_content={hc}, has_sectPr={hs}")
    print(xml_str[:400])

# Simulate the backward scan in insert_break_before_xml
print(f"\n--- Backward scan from j={tgt_idx-1} ---")
children = list(doc.element.body)
for j in range(tgt_idx - 1, max(tgt_idx - 15, -1), -1):
    elem = children[j]
    tag = elem.tag
    if not tag.endswith('}p') and not tag.endswith('}tbl') and not tag.endswith('}sdt'):
        print(f"  [{j}] tag={tag.split('}')[-1]} — NON-P/TBL")
        continue
    if tag.endswith('}p'):
        pPr = elem.find(qn('w:pPr'))
        sectPr = pPr.find(qn('w:sectPr')) if pPr is not None else None
        has_pgSz = sectPr is not None and sectPr.find(qn('w:pgSz')) is not None
        has_content = DocProcessor._p_has_content(elem)
        txt = DocProcessor._p_text(elem)[:20]
        print(f"  [{j}] p, txt={repr(txt)}, has_content={has_content}, sectPr={'yes(pgSz)' if has_pgSz else 'yes(no pgSz)' if sectPr is not None else 'no'}")
        if has_content:
            print(f"    → WOULD call _attach_sectPr([{j}])")
            break
    elif tag.endswith('}tbl') or tag.endswith('}sdt'):
        print(f"  [{j}] tbl/sdt → WOULD insert new_p before [{tgt_idx}]")
        break
