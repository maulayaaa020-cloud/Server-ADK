"""Debug sectPr XML section 1 dari output cover 1 dimulai dari ii — bandingkan OK vs problem"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

ROOT   = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
HASIL  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'hasil')
W      = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

files = ['Docx 2_c1.docx', 'Docx 4_c1.docx',   # should work
         'Docx 3_c1.docx', 'Docx 9_c1.docx', 'Docx 14_c1.docx']  # problem files

for fname in files:
    path = os.path.join(HASIL, fname)
    if not os.path.exists(path): print(f"MISSING: {fname}"); continue
    doc = Document(path)

    sec = doc.sections[1]  # roman section
    sp  = sec._sectPr
    tp  = sp.find(qn('w:titlePg'))
    pnt = sp.find(qn('w:pgNumType'))

    f      = sec.footer
    fparas = f.paragraphs
    has_fld = any(
        p.find(qn('w:fldChar')) is not None or
        p.find(qn('w:instrText')) is not None or
        p.find(qn('w:fldSimple')) is not None
        for fpar in fparas
        for p in fpar._p.iter()
    )

    ftrs = [e.get(qn('w:type')) for e in sp.findall(qn('w:footerReference'))]
    print(f"{fname:<25}  titlePg={'YES' if tp else 'NO ':3}  "
          f"start={pnt.get(qn('w:start')) if pnt is not None else '?'}  "
          f"footer_linked={f.is_linked_to_previous}  "
          f"footer_has_field={has_fld}  "
          f"footerRefs={ftrs}")

    # Dump footer paragraph XML untuk lihat field code
    for fpar in fparas:
        xml = etree.tostring(fpar._p, pretty_print=False).decode()
        print(f"  footer_para_xml: {xml[:200]}")
