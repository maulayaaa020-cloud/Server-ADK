"""Tampilkan semua content elements di cover section untuk Docx 3, 9, 14"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')))
from docx import Document
from utils import DocProcessor, is_roman_start

W    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii')

def _has_pgbr(el):
    return any(br.get('{%s}type'%W)=='page' for br in el.iter('{%s}br'%W))

def _has_sectPr(el):
    pPr = el.find('{%s}pPr'%W)
    return pPr is not None and pPr.find('{%s}sectPr'%W) is not None

def _txt(el):
    return ''.join(t.text or '' for t in el.iter('{%s}t'%W)).strip()

def _has_img(el):
    return el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline') is not None

for fname in ['Docx 3.docx', 'Docx 9.docx', 'Docx 14.docx']:
    path = os.path.join(SRC, fname)
    if not os.path.exists(path):
        print(f"MISSING: {fname}"); continue

    doc  = Document(path)
    body = list(doc.element.body)
    proc = DocProcessor(doc, 'Times New Roman', 12)
    roman_start_p, _ = proc.scan_zones()

    if roman_start_p is None:
        print(f"\n{fname}: roman_start_p=None"); continue

    try:
        rsp_i = body.index(roman_start_p)
    except ValueError:
        print(f"\n{fname}: roman_start_p not in body"); continue

    rsp_txt = _txt(roman_start_p)
    print(f"\n{'='*70}")
    print(f"{fname}  rsp_i={rsp_i}  rsp_text={repr(rsp_txt[:50])}")
    print(f"\n  SEMUA elemen [0..{rsp_i}] (cover content):")

    content_count = 0
    for i in range(rsp_i + 1):
        el = body[i]
        if not el.tag.endswith('}p'):
            print(f"  [{i:3}] <tbl/sdt>")
            continue
        txt = _txt(el)
        has_img_tag = 'IMG' if _has_img(el) else ''
        pgbr = 'pgBr' if _has_pgbr(el) else ''
        secpr = 'sectPr' if _has_sectPr(el) else ''
        is_rsp = '<-- RSP' if i == rsp_i else ''
        if txt or has_img_tag or pgbr or secpr:
            content_count += 1
            print(f"  [{i:3}] {repr(txt[:55]):58} {has_img_tag:3} {pgbr:4} {secpr:6} {is_rsp}")
        else:
            print(f"  [{i:3}] (empty)")
