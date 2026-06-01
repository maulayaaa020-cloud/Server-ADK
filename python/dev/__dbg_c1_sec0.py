"""Cek tipe sectPr dan isi elemen di sekitar boundary section 0 untuk output cover 1"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

ROOT  = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
W     = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def _has_sectPr(el):
    pPr = el.find('{%s}pPr'%W)
    return pPr is not None and pPr.find('{%s}sectPr'%W) is not None

def _has_pgbr(el):
    return any(br.get('{%s}type'%W)=='page' for br in el.iter('{%s}br'%W))

def _txt(el):
    return ''.join(t.text or '' for t in el.iter('{%s}t'%W)).strip()[:40]

for scenario, fname, folder in [
    ('OK',  'Docx 2_c1.docx', 'cover 1 dimulai dari ii'),
    ('PROB','Docx 14_c1.docx','cover 1 dimulai dari ii'),
]:
    path = os.path.join(ROOT, 'test_files', 'paket3', folder, 'hasil', fname)
    if not os.path.exists(path): print(f"MISSING: {fname}"); continue
    doc  = Document(path)
    body = list(doc.element.body)

    # Cari sectPr pertama
    sec0_end = next((i for i,el in enumerate(body) if el.tag.endswith('}p') and _has_sectPr(el)), None)
    print(f"\n[{scenario}] {fname}  sec0_end={sec0_end}")

    if sec0_end is not None:
        el = body[sec0_end]
        pPr = el.find('{%s}pPr'%W)
        sectPr = pPr.find('{%s}sectPr'%W) if pPr is not None else None
        tp = sectPr.find('{%s}type'%W) if sectPr is not None else None
        pg = sectPr.find('{%s}pgSz'%W) if sectPr is not None else None
        has_pgbr_in_boundary = _has_pgbr(el)
        print(f"  boundary para: text={repr(_txt(el))}  has_pgbr={has_pgbr_in_boundary}")
        print(f"  sectPr type={tp.get('{%s}val'%W) if tp is not None else 'None'}  has_pgSz={pg is not None}")

    # Tampilkan 6 elemen sebelum dan sesudah boundary
    if sec0_end is not None:
        start = max(0, sec0_end-4)
        end   = min(len(body), sec0_end+4)
        print(f"  context [{start}..{end-1}]:")
        for i in range(start, end):
            el = body[i]
            tag = el.tag.split('}')[-1]
            txt = _txt(el) if el.tag.endswith('}p') else ''
            sp  = 'sectPr' if _has_sectPr(el) else ''
            br  = 'pgBr' if _has_pgbr(el) else ''
            marker = ' <<<' if i == sec0_end else ''
            print(f"    [{i}] <{tag}> {repr(txt):25} {sp:6} {br}{marker}")
