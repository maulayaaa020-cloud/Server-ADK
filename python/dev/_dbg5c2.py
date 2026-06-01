import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")

from docx import Document
from docx.oxml.ns import qn
import utils as _u, paket3

# Patch fmt_roman untuk trace
_orig = _u.DocProcessor.fmt_roman
def _traced(self, section, start=None, roman_sec=None):
    sectPr = section._sectPr
    _type_el = sectPr.find(qn("w:type"))
    _is_cont = (_type_el is not None and _type_el.get(qn("w:val")) == "continuous")
    _secs = list(self.doc.sections)
    _idx = next((i for i,s in enumerate(_secs) if s._sectPr is sectPr), -1)
    print("  fmt_roman sec[%02d] continuous=%-5s start=%-4s roman_sec=%s" % (_idx, _is_cont, start, roman_sec))
    return _orig(self, section, start=start, roman_sec=roman_sec)
_u.DocProcessor.fmt_roman = _traced

SRC = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\Docx 5.docx"

doc  = Document(SRC)
proc = _u.DocProcessor(doc, "Times New Roman", 12)
proc.purge_all_headers_footers()
proc.strip_column_breaks_in_tables()

roman_start_p, bab_p_list = proc.scan_zones()
import zipfile; _use_exact = False
new_rsp, _use_exact = _u.DocProcessor.advance_roman_start(doc, roman_start_p, 2)
if new_rsp is not roman_start_p:
    roman_start_p = new_rsp
roman_start_p = proc.insert_breaks(roman_start_p, bab_p_list, exact_roman_start=_use_exact)
proc.ensure_cover_pages(roman_start_p, 2, already_advanced=_use_exact)
roman_sec, bab_sec_list, n_sections = proc.build_section_map(roman_start_p, bab_p_list)
print("roman_sec=%s  bab_sec_list=%s  n_sections=%s" % (roman_sec, bab_sec_list, n_sections))
paket3.apply(proc, roman_sec, bab_sec_list, n_sections, "Ya", "iii", 2)
