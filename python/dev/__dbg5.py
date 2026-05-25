"""Debug struktur Docx 5 untuk advance_roman_start"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
PATH = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'Docx 5.docx')

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

try:
    rsp_idx = body_els.index(roman_start_p)
except ValueError:
    print("roman_start_p not found in body_els"); exit()

breaks_before = sum(1 for el in body_els[:rsp_idx] if _has_break(el))

print(f"roman_start_p idx : {rsp_idx}")
print(f"roman_start_p text: {txt(roman_start_p)!r}")
print(f"breaks_before     : {breaks_before}")
print()

# Tampilkan konteks sekitar rsp_idx
print("Paragraf sekitar roman_start_p:")
for i in range(max(0, rsp_idx-5), min(len(body_els), rsp_idx+15)):
    el  = body_els[i]
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    br  = '[BRK]' if _has_break(el) else '     '
    mk  = ' <-- RSP' if i == rsp_idx else ''
    print(f"  [{i:3d}] {tag:<5} {br}  {txt(el)!r}{mk}")

# Scan maju dari rsp_idx untuk page break
print()
print("Forward scan dari rsp_idx:")
for j in range(rsp_idx, min(rsp_idx + 30, len(body_els))):
    el  = body_els[j]
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    has_br = _has_break(el)
    pPr = el.find("{%s}pPr" % W)
    has_spr = (pPr is not None and pPr.find("{%s}sectPr" % W) is not None)
    has_pgbr = any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
    mark = '[sectPr]' if has_spr else ('[pgBr]' if has_pgbr else '')
    print(f"  [{j:3d}] {tag:<5}  {mark:<10}  {txt(el)!r}")
