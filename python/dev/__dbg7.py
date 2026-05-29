"""Debug Docx 7 untuk cover 2"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
PATH = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'Docx 7.docx')

doc      = Document(PATH)
proc     = DocProcessor(doc, 'Times New Roman', 12)
roman_start_p, bab_p_list = proc.scan_zones()
body_els = list(doc.element.body)

def _has_break(el):
    if any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W)):
        return True
    pPr = el.find("{%s}pPr" % W)
    return pPr is not None and pPr.find("{%s}sectPr" % W) is not None

def txt(el):
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:60]

rsp_idx = body_els.index(roman_start_p)
breaks_before = sum(1 for el in body_els[:rsp_idx] if _has_break(el))

print(f"roman_start_p idx : {rsp_idx}")
print(f"roman_start_p text: {txt(roman_start_p)!r}")
print(f"breaks_before     : {breaks_before}")
print(f"bab paragraphs    : {[txt(p) for p in bab_p_list]}")
print(f"bab body indices  : {[body_els.index(p) for p in bab_p_list]}")
print()

new_rsp, use_exact = DocProcessor.advance_roman_start(doc, roman_start_p, 2)
new_rsp_idx = body_els.index(new_rsp)
print(f"advance_roman_start -> [{new_rsp_idx}] {txt(new_rsp)!r}, use_exact={use_exact}")
print()

print("Context around rsp and new_rsp:")
lo = max(0, rsp_idx - 5)
hi = min(len(body_els), new_rsp_idx + 15)
for i in range(lo, hi):
    el  = body_els[i]
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    pPr = el.find("{%s}pPr" % W)
    has_spr  = (pPr is not None and pPr.find("{%s}sectPr" % W) is not None)
    has_pgbr = any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
    mark = '[sectPr]' if has_spr else ('[pgBr]' if has_pgbr else '       ')
    mk   = ' <-- RSP' if i == rsp_idx else ''
    mk  += ' <-- NEW_RSP' if i == new_rsp_idx else ''
    mk  += ' <-- BAB' if el in bab_p_list else ''
    print(f"  [{i:3d}] {tag:<5}  {mark}  {txt(el)!r}{mk}")
