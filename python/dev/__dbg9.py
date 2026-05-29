"""Debug Docx 9 fix: count non-empty paragraphs before last_break for all bb=1 files"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from utils import DocProcessor

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii')

def _has_break(el):
    if any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W)):
        return True
    pPr = el.find("{%s}pPr" % W)
    return pPr is not None and pPr.find("{%s}sectPr" % W) is not None

def _has_sectPr(el):
    pPr = el.find("{%s}pPr" % W)
    return pPr is not None and pPr.find("{%s}sectPr" % W) is not None

def _has_pgbr(el):
    return any(br.get("{%s}type" % W) == 'page' for br in el.iter("{%s}br" % W))

def txt(el, n=35):
    return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:n]

def el_tag(el):
    return el.tag.split('}')[-1] if '}' in el.tag else el.tag

print(f"{'N':>4}  {'bb':>3}  {'rsp_idx':>7}  {'last_brk':>8}  {'brk_type':12}  {'nonempty':>8}  {'advance?':15}  rsp_txt")
print("-" * 100)

for n in range(1, 19):
    fname = f"Docx {n}.docx"
    path  = os.path.join(SRC, fname)
    if not os.path.exists(path):
        continue

    doc      = Document(path)
    proc     = DocProcessor(doc, 'Times New Roman', 12)
    rsp, bab = proc.scan_zones()
    body     = list(doc.element.body)
    rsp_idx  = body.index(rsp)
    bb       = sum(1 for el in body[:rsp_idx] if _has_break(el))

    # Find last break before rsp
    last_brk_idx  = -1
    last_brk_type = '(none)'
    for j in range(rsp_idx - 1, -1, -1):
        if _has_break(body[j]):
            last_brk_idx = j
            if _has_sectPr(body[j]) and _has_pgbr(body[j]):
                last_brk_type = 'sectPr+pgBr'
            elif _has_sectPr(body[j]):
                last_brk_type = 'sectPr:' + ('empty' if not txt(body[j]) else f'"{txt(body[j], 12)}"')
            else:
                last_brk_type = 'pgBr:' + ('empty' if not txt(body[j]) else f'"{txt(body[j], 12)}"')
            break

    # Count non-empty 'p' paragraphs before last_break
    nonempty_before = sum(
        1 for i in range(last_brk_idx)
        if el_tag(body[i]) == 'p' and txt(body[i])
    )

    new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 2)
    new_idx = body.index(new_rsp)
    advance_info = f"[{rsp_idx}]->[{new_idx}]" if new_rsp is not rsp else f"[{rsp_idx}] (no chg)"

    print(f"  {n:>2}  {bb:>3}  {rsp_idx:>7}  {last_brk_idx:>8}  {last_brk_type:<12}  {nonempty_before:>8}  {advance_info:<15}  {txt(rsp)!r}")

print()
print("Focus: bb=1 files and their nonempty count before last_break")
