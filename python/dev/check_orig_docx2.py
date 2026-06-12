"""Cek struktur Docx 2 ORIGINAL (sebelum diproses) dan HASIL secara paralel."""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import zipfile
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
SEP = "=" * 70

def check_file(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    # Cari SDT
    sdts = [c for c in children if c.tag == f"{W}sdt"]
    # Cari instrText TOC
    instrs = body.findall(f".//{W}instrText")
    toc_instrs = [i for i in instrs if i.text and "TOC" in i.text]

    print(f"\n{SEP}")
    print(f" {label}: {len(children)} paragraf, {len(sdts)} SDT, {len(toc_instrs)} TOC instrText")
    print(SEP)

    # Cari heading DAFTAR PUSTAKA
    print("\nHeading sekitar DAFTAR PUSTAKA / DAFTAR ISI:")
    for i, child in enumerate(children):
        pPr = child.find(f"{W}pPr")
        pstyle = None
        if pPr is not None:
            ps = pPr.find(f"{W}pStyle")
            if ps is not None:
                pstyle = ps.get(f"{W}val")
        txt = "".join((t.text or "") for t in child.findall(f".//{W}t")).strip()
        if "DAFTAR" in txt.upper() or "PUSTAKA" in txt.upper():
            print(f"  [{i}] style={str(pstyle):15} | {txt[:50]!r}")

# Original (sebelum diproses)
check_file(r"D:\Freelaces\Test Dafis\Docx 2.docx", "DOCX 2 ORIGINAL")

# Hasil
check_file(r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx", "DOCX 2 HASIL")

# File Benar
check_file(r"D:\Freelaces\Test Dafis\File Benar\Docx 2.docx", "DOCX 2 FILE BENAR")
