"""Lihat konten SEBELUM DAFTAR ISI di file yang perlu dianalisa."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

def scan_before(path, fname):
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
        print(f"{fname}: DAFTAR ISI tidak ditemukan"); return

    print(f"\n=== {fname} ORIGINAL: total {len(paras)} paras, DAFTAR ISI=[{di}] ===")
    print("--- 15 PARA SEBELUM DAFTAR ISI ---")
    for i in range(max(0, di-15), di):
        p = paras[i]
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        flds  = p.findall(f".//{W}fldChar")
        brs   = p.findall(f".//{W}br")
        pPr   = p.find(f"{W}pPr")
        sn    = pPr.find(f"{W}pStyle") if pPr is not None else None
        style = sn.get(f"{W}val","") if sn is not None else ""
        sp    = pPr.find(f"{W}sectPr") if pPr is not None else None
        flags = []
        if flds:  flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
        if brs:   flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        if sp is not None:
            t2 = sp.find(f"{W}type")
            flags.append("sect(" + (t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage") + ")")
        print(f"  [{i:3}] style={style:12} | {repr(text[:50]):55} | {' '.join(flags)}")

    print("--- [DI] = DAFTAR ISI ---")
    print("--- 20 PARA SETELAH DAFTAR ISI ---")
    for i in range(di, min(di+20, len(paras))):
        p = paras[i]
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        flds  = p.findall(f".//{W}fldChar")
        brs   = p.findall(f".//{W}br")
        pPr   = p.find(f"{W}pPr")
        sn    = pPr.find(f"{W}pStyle") if pPr is not None else None
        style = sn.get(f"{W}val","") if sn is not None else ""
        sp    = pPr.find(f"{W}sectPr") if pPr is not None else None
        flags = []
        if flds:  flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
        if brs:   flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        if sp is not None:
            t2 = sp.find(f"{W}type")
            flags.append("sect(" + (t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage") + ")")
        print(f"  [{i:3}] style={style:12} | {repr(text[:50]):55} | {' '.join(flags)}")

ORIG = r"D:\Freelaces\Test Dafis"
for fname in ["Docx 6.docx", "Docx 14.docx", "Docx 15.docx", "Docx 17.docx",
              "Docx 10.docx", "Docx 2.docx", "Docx 3.docx"]:
    scan_before(os.path.join(ORIG, fname), fname)
