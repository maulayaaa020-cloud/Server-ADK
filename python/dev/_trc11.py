import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document
from docx.oxml.ns import qn
import utils as _u, paket3

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def has_pgbr(el):
    return any(br.get("{%s}type"%W)=="page" for br in el.iter("{%s}br"%W))
def has_sectpr(el):
    pPr = el.find("{%s}pPr"%W)
    return pPr is not None and pPr.find("{%s}sectPr"%W) is not None
def txt(el):
    return "".join(t.text or "" for t in el.iter("{%s}t"%W)).strip()[:50]

SRC = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\Docx 11.docx"
doc = Document(SRC)
body_els = list(doc.element.body)

from utils import DocProcessor
proc = DocProcessor(doc, "Times New Roman", 12)
roman_start_p, bab_p_list = proc.scan_zones()
rsp_idx = body_els.index(roman_start_p)

print("=== ORIGINAL roman_start_p ===")
print("  idx=%d  txt=%r" % (rsp_idx, txt(roman_start_p)))

new_rsp, use_exact = DocProcessor.advance_roman_start(doc, roman_start_p, 2)
new_idx = body_els.index(new_rsp)
print("\n=== AFTER advance_roman_start(num_cover=2) ===")
print("  use_exact=%s  new_idx=%d  txt=%r" % (use_exact, new_idx, txt(new_rsp)))

print("\n=== Page breaks in cover zone [0..%d] ===" % new_idx)
for j in range(new_idx):
    el = body_els[j]
    if el.tag.endswith("}p") and has_pgbr(el):
        print("  [%03d] PGBR  %r" % (j, txt(el)))
    elif el.tag.endswith("}p") and has_sectpr(el):
        print("  [%03d] SECT  %r" % (j, txt(el)))
