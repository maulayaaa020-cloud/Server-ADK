"""Scan area DAFTAR ISI di File Benar untuk semua 20 file."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

folder = r"D:\Freelaces\Test Dafis\File Benar"

for fname in sorted(os.listdir(folder)):
    if not fname.endswith(".docx"):
        continue
    path = os.path.join(folder, fname)
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")
    paras = body.findall(f"{W}p")

    # Cari DAFTAR ISI heading
    di = None
    for i, p in enumerate(paras):
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip().upper()
        if txt == "DAFTAR ISI":
            di = i
            break

    if di is None:
        print(f"{fname}: DAFTAR ISI tidak ditemukan")
        continue

    # Dump 15 paragraf dari DAFTAR ISI
    print(f"\n=== {fname} (DAFTAR ISI=[{di}]) ===")
    for i in range(di, min(di+15, len(paras))):
        p = paras[i]
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        flds  = p.findall(f".//{W}fldChar")
        instrs= p.findall(f".//{W}instrText")
        brs   = p.findall(f".//{W}br")
        pPr   = p.find(f"{W}pPr")
        sn    = pPr.find(f"{W}pStyle") if pPr is not None else None
        style = sn.get(f"{W}val","") if sn is not None else ""
        sp    = pPr.find(f"{W}sectPr") if pPr is not None else None
        flags = []
        if flds:   flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
        if instrs: flags.append("instr:" + (instrs[0].text or "")[:40].strip())
        if brs:    flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        if sp is not None:
            t2 = sp.find(f"{W}type")
            flags.append("sect(" + (t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage") + ")")
        print(f"  [{i:3}] style={style:18} | {repr(text[:55]):58} | {' '.join(flags)}")
