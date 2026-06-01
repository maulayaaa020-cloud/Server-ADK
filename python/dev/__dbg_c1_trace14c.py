"""Trace sangat detail insert_break_before_xml untuk Docx 14"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from utils import DocProcessor, is_roman_start
from docx.oxml.ns import qn
from copy import deepcopy
from lxml import etree

W    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'Docx 14.docx')

doc  = Document(SRC)
proc = DocProcessor(doc, 'Times New Roman', 12)
rsp, bab_p_list = proc.scan_zones()

body = list(doc.element.body)
rsp_i = body.index(rsp)
print(f"RSP at [{rsp_i}] = '{DocProcessor._p_text(rsp)[:30]}'")

# Advance
new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
new_body = list(doc.element.body)
new_rsp_i = new_body.index(new_rsp) if new_rsp in new_body else -1
print(f"After advance: new_rsp at [{new_rsp_i}]")

def _has_sectPr(el):
    pPr = el.find('{%s}pPr'%W)
    return pPr is not None and pPr.find('{%s}sectPr'%W) is not None

def _show_sec(label, b):
    secs = [(i, repr(DocProcessor._p_text(e)[:20])) for i,e in enumerate(b[:55]) if e.tag.endswith('}p') and _has_sectPr(e)]
    print(f"  {label} sectPr in [0..54]: {secs}")

_show_sec("after advance", new_body)

# Manually call _find_section_start
fss = proc._find_section_start(new_rsp)
fss_i = list(doc.element.body).index(fss) if fss in list(doc.element.body) else -1
print(f"_find_section_start result: [{fss_i}]")
_show_sec("after fss", list(doc.element.body))

# _roman_already_bounded check
body_ch = list(doc.element.body)
_rsp_idx = next((i for i,e in enumerate(body_ch) if e is fss), -1)
rab = (
    _rsp_idx > 0 and
    body_ch[_rsp_idx - 1].tag.endswith('}p') and
    proc._has_sectPr(body_ch[_rsp_idx - 1])
)
print(f"_rsp_idx={_rsp_idx}, _roman_already_bounded={rab}")
print(f"  body_ch[{_rsp_idx-1}] text: '{DocProcessor._p_text(body_ch[_rsp_idx-1])[:30]}'")
print(f"  body_ch[{_rsp_idx-1}] has_sectPr: {proc._has_sectPr(body_ch[_rsp_idx-1])}")

# Now call insert_break_before_xml
print(f"\nCalling insert_break_before_xml([{_rsp_idx}])...")
proc.insert_break_before_xml(fss)
_show_sec("after insert_break_before_xml", list(doc.element.body))

# _strip_empty_paras_before_bab
print(f"\nCalling _strip_empty_paras_before_bab([{_rsp_idx}])...")
proc._strip_empty_paras_before_bab(fss)
after_strip = list(doc.element.body)
_show_sec("after strip", after_strip)
print(f"  total body elements: {len(after_strip)}")
# Show around where sectPr is
for sec_i, sec_txt in [(i, repr(DocProcessor._p_text(e)[:30])) for i,e in enumerate(after_strip[:55]) if e.tag.endswith('}p') and _has_sectPr(e)]:
    start = max(0, sec_i - 3)
    end   = min(len(after_strip), sec_i + 5)
    print(f"\n  sectPr at [{sec_i}], context [{start}..{end-1}]:")
    for j in range(start, end):
        e = after_strip[j]
        txt = repr(DocProcessor._p_text(e)[:30]) if e.tag.endswith('}p') else '<tbl>'
        sp = 'sectPr' if _has_sectPr(e) else ''
        m = ' <<<' if j == sec_i else ''
        print(f"    [{j}] {txt:35} {sp}{m}")
