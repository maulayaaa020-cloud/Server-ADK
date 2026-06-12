"""Lihat DAFTAR ISI di Original vs File Benar untuk memahami perbedaan."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
ORIG  = r"D:\Freelaces\Test Dafis"
BENAR = r"D:\Freelaces\Test Dafis\File Benar"
HASIL = r"D:\Freelaces\Test Dafis\Hasil"

def scan(path, label, n=12):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")
    paras = body.findall(f"{W}p")

    di = None
    for i, p in enumerate(paras):
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip().upper()
        if txt == "DAFTAR ISI":
            di = i
            break
    if di is None:
        print(f"  {label}: DAFTAR ISI tidak ditemukan")
        return

    print(f"  {label} [{di}]:", end="")
    items = []
    for j in range(di+1, min(di+n, len(paras))):
        p = paras[j]
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        flds  = p.findall(f".//{W}fldChar")
        pPr   = p.find(f"{W}pPr")
        sn    = pPr.find(f"{W}pStyle") if pPr is not None else None
        style = sn.get(f"{W}val","") if sn is not None else ""
        if text or flds:
            items.append(f"{style}/{repr(text[:30])}")
    print(" | ".join(items[:5]) if items else " (kosong)")

print("=== PERBANDINGAN ORIG vs BENAR vs HASIL ===\n")
for fname in sorted(os.listdir(ORIG)):
    if not fname.lower().endswith(".docx") or not fname.startswith("Docx"): continue
    op = os.path.join(ORIG,  fname)
    bp = os.path.join(BENAR, fname)
    hp = os.path.join(HASIL, fname)
    print(f"\n{fname}:")
    scan(op, "ORIG ")
    if os.path.exists(bp): scan(bp, "BENAR")
    if os.path.exists(hp): scan(hp, "HASIL")
