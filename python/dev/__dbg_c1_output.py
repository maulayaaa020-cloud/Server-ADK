"""Debug struktur output file Docx 3,9,14 — cek sec0 boundary dan halaman"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

ROOT  = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
HASIL = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'hasil')
W     = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def _has_sectPr(el):
    pPr = el.find('{%s}pPr'%W)
    return pPr is not None and pPr.find('{%s}sectPr'%W) is not None

def _has_pgbr(el):
    return any(br.get('{%s}type'%W)=='page' for br in el.iter('{%s}br'%W))

def _txt(el):
    return ''.join(t.text or '' for t in el.iter('{%s}t'%W)).strip()[:50]

def _has_content(el):
    return bool(_txt(el)) or any(el.iter('{%s}drawing'%W)) or any(el.iter('{%s}pict'%W))

for fname in ['Docx 3_c1.docx', 'Docx 9_c1.docx', 'Docx 14_c1.docx']:
    path = os.path.join(HASIL, fname)
    if not os.path.exists(path):
        print(f"MISSING: {fname}"); continue

    doc  = Document(path)
    body = list(doc.element.body)

    # Cari semua sectPr
    sec_positions = [(i, el) for i, el in enumerate(body)
                     if el.tag.endswith('}p') and _has_sectPr(el)]
    # sectPr terakhir ada di body langsung
    final_sectPr = doc.element.body.find('{%s}sectPr'%W)

    print(f"\n{'='*70}")
    print(f"{fname}  total_body={len(body)}")
    print(f"  sectPr di paragraph: {[(i, repr(_txt(el))) for i,el in sec_positions]}")
    print(f"  final sectPr (doc level): {'YES' if final_sectPr is not None else 'NO'}")

    # Cek section 0: semua elemen sampai sectPr pertama
    if sec_positions:
        sec0_end, sec0_el = sec_positions[0]
        sectPr = sec0_el.find('{%s}pPr'%W).find('{%s}sectPr'%W)
        tp = sectPr.find('{%s}type'%W)
        pg = sectPr.find('{%s}pgSz'%W)
        pnt = sectPr.find('{%s}pgNumType'%W)
        print(f"\n  Section 0 ends at [{sec0_end}]:")
        print(f"    text={repr(_txt(sec0_el))}")
        print(f"    type={tp.get('{%s}val'%W) if tp is not None else 'None'}")
        print(f"    has_pgSz={pg is not None}")
        print(f"    pgNumType start={pnt.get('{%s}start'%W) if pnt is not None else 'None'}")

        # Tampilkan context sekitar sec0_end
        start = max(0, sec0_end - 5)
        end   = min(len(body), sec0_end + 5)
        print(f"\n  Context [{start}..{end-1}]:")
        for i in range(start, end):
            el = body[i]
            tag = el.tag.split('}')[-1]
            txt = _txt(el) if el.tag.endswith('}p') else ''
            has_img = 'IMG' if any(el.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')) else ''
            sp  = 'sectPr' if _has_sectPr(el) else ''
            br  = 'pgBr' if _has_pgbr(el) else ''
            cnt = 'CONTENT' if _has_content(el) else 'empty'
            marker = ' <<<' if i == sec0_end else ''
            print(f"    [{i:3}] <{tag}> {repr(txt):35} {cnt:7} {sp:6} {br}{has_img}{marker}")

    # Hitung elemen dengan content di section 0
    if sec_positions:
        content_count = sum(1 for i in range(sec_positions[0][0]+1) if _has_content(body[i]))
        pgbr_count    = sum(1 for i in range(sec_positions[0][0]+1) if _has_pgbr(body[i]))
        print(f"\n  Section 0: {sec_positions[0][0]+1} elements, {content_count} with content, {pgbr_count} pgBr")

    # Cek section 1 pgNumType
    if len(sec_positions) >= 2:
        sec1_end, sec1_el = sec_positions[1]
        sectPr1 = sec1_el.find('{%s}pPr'%W).find('{%s}sectPr'%W)
        pnt1 = sectPr1.find('{%s}pgNumType'%W)
        tp1 = sectPr1.find('{%s}type'%W)
        print(f"\n  Section 1 ends at [{sec1_end}] text={repr(_txt(sec1_el))}")
        print(f"    type={tp1.get('{%s}val'%W) if tp1 is not None else 'None'}")
        print(f"    pgNumType start={pnt1.get('{%s}start'%W) if pnt1 is not None else 'None'}")
