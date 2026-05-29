"""Debug Docx 14 - full break/section structure"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
PATH = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'Docx 14.docx')

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

new_rsp, use_exact = DocProcessor.advance_roman_start(doc, roman_start_p, 2)
new_rsp_idx = body_els.index(new_rsp)

print(f"rsp={rsp_idx} ({txt(roman_start_p)!r}), new_rsp={new_rsp_idx} ({txt(new_rsp)!r}), bb={breaks_before}")
print(f"bab[0] at body[{bab_indices[0]}]: {txt(bab_p_list[0])!r}")
print()

# Show context around new_rsp (wider)
print("Context [new_rsp-5 .. new_rsp+10]:")
for i in range(max(0, new_rsp_idx-5), min(len(body_els), new_rsp_idx+11)):
    el  = body_els[i]
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    pPr = el.find("{%s}pPr" % W)
    has_spr  = (pPr is not None and pPr.find("{%s}sectPr" % W) is not None)
    has_pgbr = any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
    mark = '[sectPr]' if has_spr else ('[pgBr]' if has_pgbr else '       ')
    mk   = ' <-- RSP' if i == rsp_idx else ''
    mk  += ' <-- NEW_RSP' if i == new_rsp_idx else ''
    mk  += ' <-- BAB' if i in bab_indices else ''
    print(f"  [{i:3d}] {tag:<5}  {mark}  {txt(el)!r}{mk}")

print()
print(f"Context [bab[0]-8 .. bab[0]+3]:")
bi0 = bab_indices[0]
for i in range(max(0, bi0-8), min(len(body_els), bi0+4)):
    el  = body_els[i]
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    pPr = el.find("{%s}pPr" % W)
    has_spr  = (pPr is not None and pPr.find("{%s}sectPr" % W) is not None)
    has_pgbr = any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
    mark = '[sectPr]' if has_spr else ('[pgBr]' if has_pgbr else '       ')
    mk   = ' <-- BAB' if i in bab_indices else ''
    has_content = proc._p_has_content(el)
    cts = 'C' if has_content else ' '
    print(f"  [{i:3d}] {cts}{tag:<5}  {mark}  {txt(el)!r}{mk}")

# Check the output file too
HASIL = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'hasil', 'Docx 14_v2.docx')
if os.path.exists(HASIL):
    doc2 = Document(HASIL)
    print()
    print("OUTPUT sections:")
    for i, sec in enumerate(doc2.sections):
        pn = sec._sectPr.find(qn('w:pgNumType'))
        fmt   = pn.get(qn('w:fmt'), 'decimal') if pn is not None else 'decimal'
        start = pn.get(qn('w:start'), None) if pn is not None else None
        print(f"  sec[{i}]: {fmt}:{start}, diff={sec.different_first_page_header_footer}")
