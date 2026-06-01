"""Debug cover 1: lihat struktur elemen antara page break (bb) dan roman_start_p untuk Docx 3,9,14"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor, is_roman_start

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ROOT    = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
FOLDER  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii')

def _is_page_break(el):
    if any(br.get('{%s}type' % W) == 'page' for br in el.iter('{%s}br' % W)):
        return True
    pPr = el.find('{%s}pPr' % W)
    return pPr is not None and pPr.find('{%s}sectPr' % W) is not None

for fname in ['Docx 3.docx', 'Docx 9.docx', 'Docx 14.docx']:
    path = os.path.join(FOLDER, fname)
    if not os.path.exists(path):
        print(f"  MISSING: {fname}"); continue

    doc  = Document(path)
    font = 'Times New Roman'
    proc = DocProcessor(doc, font, 12)
    roman_start_p, _ = proc.scan_zones()

    body_els = list(doc.element.body)

    if roman_start_p is None:
        print(f"\n{fname}: roman_start_p=None"); continue

    try:
        rsp_i = body_els.index(roman_start_p)
    except ValueError:
        print(f"\n{fname}: roman_start_p not in body"); continue

    # Hitung bb dan posisi tiap break
    bb_positions = []
    for i, el in enumerate(body_els[:rsp_i]):
        if _is_page_break(el):
            bb_positions.append(i)

    print(f"\n{'='*60}")
    print(f"{fname}  rsp_i={rsp_i}  bb={len(bb_positions)}  pgBrs@{bb_positions}")
    print(f"  roman_start_p text: {repr(DocProcessor._p_text(roman_start_p)[:60])}")

    # Tampilkan elemen setelah BREAK TERAKHIR sampai roman_start_p
    if bb_positions:
        last_bb = bb_positions[-1]
        print(f"\n  Elemen setelah break terakhir (idx {last_bb}) sampai rsp (idx {rsp_i}):")
        for i in range(last_bb + 1, rsp_i + 1):
            el = body_els[i]
            tag  = el.tag.split('}')[-1]
            text = DocProcessor._p_text(el)[:50] if el.tag.endswith('}p') else ''
            has_br = _is_page_break(el)
            marker = ' <-- roman_start_p' if i == rsp_i else ''
            print(f"    [{i}] <{tag}> {repr(text)}{' [PAGE_BREAK]' if has_br else ''}{marker}")
    else:
        print(f"  Tidak ada break sebelum roman_start_p (bb=0)")
        print(f"\n  5 elemen sebelum roman_start_p:")
        for i in range(max(0, rsp_i-5), rsp_i+1):
            el = body_els[i]
            tag  = el.tag.split('}')[-1]
            text = DocProcessor._p_text(el)[:50] if el.tag.endswith('}p') else ''
            marker = ' <-- roman_start_p' if i == rsp_i else ''
            print(f"    [{i}] <{tag}> {repr(text)}{marker}")
