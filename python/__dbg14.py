"""Debug Docx 14 cover 2"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

for fname, num_cover in [('Docx 14.docx', 2)]:
    PATH = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', fname)
    doc  = Document(PATH)
    proc = DocProcessor(doc, 'Times New Roman', 12)
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

    new_rsp, use_exact = DocProcessor.advance_roman_start(doc, roman_start_p, num_cover)
    new_rsp_idx = body_els.index(new_rsp)

    print(f"{fname}: rsp_idx={rsp_idx}, bb={breaks_before}, new_rsp_idx={new_rsp_idx}, use_exact={use_exact}")
    print(f"  rsp text: {txt(roman_start_p)!r}")
    print(f"  new_rsp text: {txt(new_rsp)!r}")
    print(f"  bab indices: {bab_indices}")
    print(f"  bab texts: {[txt(p) for p in bab_p_list]}")
    print()

    # Show context around new_rsp
    lo = max(0, rsp_idx - 3)
    hi = min(len(body_els), new_rsp_idx + 10)
    for i in range(lo, hi):
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
