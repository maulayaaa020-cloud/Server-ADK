"""Lihat struktur DAFTAR ISI di File Benar untuk beberapa file sample."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
BENAR = r"D:\Freelaces\Test Dafis\File Benar"
HASIL = r"D:\Freelaces\Test Dafis\Hasil"

def scan(path, label, n=25):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")
    paras = body.findall(f"{W}p")

    # Cari semua paras yang mengandung "DAFTAR ISI"
    hits = []
    for i, p in enumerate(paras):
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if "DAFTAR ISI" in txt.upper():
            hits.append((i, txt))

    print(f"\n{'='*70}")
    print(f"{label} | total={len(paras)} para | hits 'DAFTAR ISI': {hits[:5]}")

    # Cari yang EXACT "DAFTAR ISI"
    di = None
    for i, p in enumerate(paras):
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip().upper()
        if txt == "DAFTAR ISI":
            di = i
            break

    if di is None:
        print("  => Tidak ada exact match 'DAFTAR ISI'")
        # tampilkan semua hits
        for i, txt in hits[:3]:
            p = paras[i]
            pPr = p.find(f"{W}pPr")
            sn  = pPr.find(f"{W}pStyle") if pPr is not None else None
            style = sn.get(f"{W}val","") if sn is not None else ""
            flds  = p.findall(f".//{W}fldChar")
            flags = "fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds) if flds else ""
            print(f"  [{i}] style={style} | {repr(txt[:60])} | {flags}")
        return

    print(f"  => DAFTAR ISI at [{di}]")
    for j in range(di, min(di + n, len(paras))):
        p = paras[j]
        text  = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        flds  = p.findall(f".//{W}fldChar")
        instrs= p.findall(f".//{W}instrText")
        brs   = p.findall(f".//{W}br")
        pPr   = p.find(f"{W}pPr")
        sn    = pPr.find(f"{W}pStyle") if pPr is not None else None
        style = sn.get(f"{W}val","") if sn is not None else ""
        flags = []
        if flds:   flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
        if instrs: flags.append("instr:" + (instrs[0].text or "")[:25].strip())
        if brs:    flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        print(f"  [{j-di:3}] style={style:14} | {repr(text[:45]):50} | {' '.join(flags)[:40]}")

# Sample 4 file
for fname in ["Docx 1.docx", "Docx 3.docx", "Docx 10.docx", "Docx 12.docx"]:
    scan(os.path.join(BENAR, fname), f"BENAR/{fname}")
    scan(os.path.join(HASIL, fname), f"HASIL/{fname}")
