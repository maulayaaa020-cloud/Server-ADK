"""Check Docx 7 output - count pages in each section"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn

ROOT  = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
HASIL = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'hasil')
W     = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

for fname in ['Docx 6_v2.docx', 'Docx 7_v2.docx']:
    path = os.path.join(HASIL, fname)
    if not os.path.exists(path):
        print(f"{fname}: NOT FOUND"); continue
    doc = Document(path)
    body_els = list(doc.element.body)

    def _has_pgbr(el):
        return any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))
    def _has_spr(el):
        pPr = el.find("{%s}pPr" % W)
        return pPr is not None and pPr.find("{%s}sectPr" % W) is not None
    def txt(el):
        return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:50]

    def get_pn(sec):
        pn = sec._sectPr.find(qn('w:pgNumType'))
        if pn is None: return 'decimal:None'
        fmt   = pn.get(qn('w:fmt'), 'decimal')
        start = pn.get(qn('w:start'), None)
        return f"{fmt}:{start}"

    # Find section boundaries in body
    breaks = []  # (idx, type) where type = 'sectPr' or 'pgBr'
    for i, el in enumerate(body_els):
        if el.tag.endswith('}p'):
            if _has_spr(el):
                breaks.append((i, 'sectPr', txt(el)))
            elif _has_pgbr(el):
                breaks.append((i, 'pgBr', txt(el)))

    print(f"\n{fname}:")
    print(f"  Sections: {len(doc.sections)}")
    for i, sec in enumerate(doc.sections):
        print(f"  sec[{i}]: {get_pn(sec)}, diff_first={sec.different_first_page_header_footer}")

    print(f"\n  Breaks in body:")
    for idx, btype, t in breaks:
        print(f"    [{idx:3d}] {btype}: {t!r}")

    # Estimate pages per section
    # Count pgBr and sectPr within each section
    print(f"\n  Page estimate per section:")
    sec_boundaries = []
    for idx, btype, _ in breaks:
        if btype == 'sectPr':
            sec_boundaries.append(idx)

    prev = 0
    for si, sbnd in enumerate(sec_boundaries):
        # Count pgBr between prev and sbnd
        n_pgbr = sum(1 for idx, btype, _ in breaks if btype == 'pgBr' and prev <= idx < sbnd)
        pages_in_sec = 1 + n_pgbr
        print(f"    sec[{si}]: ~{pages_in_sec} page(s), body[{prev}..{sbnd}]")
        prev = sbnd + 1
    # Last section
    n_pgbr = sum(1 for idx, btype, _ in breaks if btype == 'pgBr' and prev <= idx)
    print(f"    sec[{len(sec_boundaries)}]: ~{1+n_pgbr} page(s), body[{prev}..]")
