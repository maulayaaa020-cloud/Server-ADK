"""Debug Docx 6 - full bab positions and page break location"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
PATH = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'Docx 6.docx')

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
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:50]

rsp_idx       = body_els.index(roman_start_p)
bab_indices   = [body_els.index(p) for p in bab_p_list]
breaks_before = sum(1 for el in body_els[:rsp_idx] if _has_break(el))

print(f"roman_start_p idx : {rsp_idx}")
print(f"breaks_before     : {breaks_before}")
print(f"bab body indices  : {bab_indices}")
print(f"bab texts         : {[txt(p) for p in bab_p_list]}")
print(f"total body_els    : {len(body_els)}")
print()

# Show all page/section breaks in document
print("All breaks in document:")
for i, el in enumerate(body_els):
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    pPr = el.find("{%s}pPr" % W)
    has_spr  = (pPr is not None and pPr.find("{%s}sectPr" % W) is not None)
    has_pgbr = any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
    if has_spr or has_pgbr:
        mark = '[sectPr]' if has_spr else '[pgBr]'
        mk   = ' <-- RSP' if i == rsp_idx else ''
        mk  += ' <-- BAB' if i in bab_indices else ''
        print(f"  [{i:3d}] {mark}  {txt(el)!r}{mk}")

print()
# Show context around first page break after rsp
first_br_after_rsp = next((i for i in range(rsp_idx, len(body_els)) if _has_break(body_els[i])), None)
if first_br_after_rsp is not None:
    print(f"First break after rsp: [{first_br_after_rsp}] dist={first_br_after_rsp-rsp_idx}")
    for i in range(max(0,first_br_after_rsp-3), min(len(body_els),first_br_after_rsp+5)):
        el  = body_els[i]
        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        pPr = el.find("{%s}pPr" % W)
        has_spr  = (pPr is not None and pPr.find("{%s}sectPr" % W) is not None)
        has_pgbr = any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
        mark = '[sectPr]' if has_spr else ('[pgBr]' if has_pgbr else '       ')
        mk   = ' <-- BAB' if i in bab_indices else ''
        print(f"  [{i:3d}] {tag:<5}  {mark}  {txt(el)!r}{mk}")
