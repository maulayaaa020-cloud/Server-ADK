"""Scan SEMUA kemunculan 'DAFTAR ISI' di original, dan semua konten antara daftar isi dan BAB I."""
import zipfile
from lxml import etree

for fname, folder in [
    ("Docx 1.docx", r"D:\Freelaces\Test Dafis"),
]:
    path = f"{folder}\\{fname}"
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W  = f"{{{ns}}}"

    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")
    paras = body.findall(f"{W}p")

    print(f"=== {fname} ({len(paras)} paras) ===\n")

    # Cari SEMUA 'DAFTAR ISI' (exact atau partial)
    print("--- Semua paragraf mengandung 'DAFTAR' ---")
    for i, p in enumerate(paras):
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if "DAFTAR" in text.upper():
            pPr = p.find(f"{W}pPr")
            sn  = pPr.find(f"{W}pStyle") if pPr is not None else None
            style = sn.get(f"{W}val","") if sn is not None else ""
            flds = p.findall(f".//{W}fldChar")
            print(f"  [{i:3}] style={style:20} | {repr(text[:70])}")

    print()

    # Cari juga paragraf antara DAFTAR ISI heading dan BAB I
    di = None
    bab1 = None
    for i, p in enumerate(paras):
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if text.upper() == "DAFTAR ISI" and di is None:
            di = i
        if text.startswith("BAB I") and di is not None and bab1 is None:
            bab1 = i
            break

    print(f"DAFTAR ISI heading at [{di}], BAB I at [{bab1}]")
    print(f"\n--- Semua paragraf antara [{di}] dan [{bab1}] ---")
    for i in range(di, bab1 + 1):
        p = paras[i]
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        flds = p.findall(f".//{W}fldChar")
        instrs = p.findall(f".//{W}instrText")
        brs  = p.findall(f".//{W}br")
        pPr  = p.find(f"{W}pPr")
        sn   = pPr.find(f"{W}pStyle") if pPr is not None else None
        style = sn.get(f"{W}val","") if sn is not None else ""
        flags = []
        if flds: flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
        if instrs: flags.append("instr:" + (instrs[0].text or "")[:30])
        if brs: flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        print(f"  [{i:3}] style={style:20} | {repr(text[:60]):65} | {' '.join(flags)}")