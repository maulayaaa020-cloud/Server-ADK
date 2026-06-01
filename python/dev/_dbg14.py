import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx.oxml.ns import qn
from utils import DocProcessor

path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 14.docx"

proc = DocProcessor(path, "Times New Roman", "12 pt")
proc.purge_all_headers_footers()
proc.insert_breaks()
proc.ensure_cover_pages(num_cover=1, dimulai_dari="ii", hidden_cov="Ya")

roman_sec, bab_sec_list, n_sections = proc.build_section_map("ii")
print("roman_sec=%s  n_sections=%s  bab_sec_list=%s" % (roman_sec, n_sections, bab_sec_list))
first_bab_sec = bab_sec_list[0] if bab_sec_list else n_sections

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
for i, sec in enumerate(proc.doc.sections):
    sp = sec._sectPr
    tp = sp.find("{%s}type" % W)
    stype = tp.get("{%s}val" % W) if tp is not None else "nextPage"
    pn = sp.find("{%s}pgNumType" % W)
    pn_attr = {k.split("}")[1]: v for k, v in pn.attrib.items()} if pn is not None else {}
    zone = "cover" if i < roman_sec else ("roman" if i < first_bab_sec else "bab")
    print("  sec[%02d] zone=%-6s type=%-14s pgNum=%s" % (i, zone, stype, pn_attr))
