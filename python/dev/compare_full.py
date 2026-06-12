"""Bandingkan SEMUA konten antara DAFTAR ISI dan konten pertama (BAB/dll) di Original vs File Benar."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
ORIG  = r"D:\Freelaces\Test Dafis"
BENAR = r"D:\Freelaces\Test Dafis\File Benar"

def scan_all(path, limit=60):
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
        return None, []

    result = []
    for i in range(di, min(di + limit, len(paras))):
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
        if instrs: flags.append("instr:" + (instrs[0].text or "")[:25].strip())
        if brs:    flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        if sp is not None:
            t2 = sp.find(f"{W}type")
            flags.append("sect(" + (t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage") + ")")
        result.append((i - di, style[:12], text[:40], " ".join(flags)[:35]))
    return di, result

# Hanya tampilkan file yang berbeda
targets = ["Docx 6.docx", "Docx 10.docx", "Docx 14.docx", "Docx 15.docx",
           "Docx 16.docx", "Docx 17.docx", "Docx 18.docx", "Docx 19.docx",
           "Docx 2.docx",  "Docx 3.docx",  "Docx 12.docx", "Docx 20.docx"]

for fname in targets:
    orig_path  = os.path.join(ORIG, fname)
    benar_path = os.path.join(BENAR, fname)
    if not os.path.exists(orig_path) or not os.path.exists(benar_path):
        continue

    di_o, rows_o = scan_all(orig_path)
    di_b, rows_b = scan_all(benar_path)

    print(f"\n{'='*80}")
    print(f"FILE: {fname}  | ORIGINAL di=[{di_o}]  |  FILE BENAR di=[{di_b}]")
    print(f"{'OFF':4} {'ORIG style':12} {'ORIG text':40} | {'BENAR style':12} {'BENAR text':40}")
    print("-"*80)

    maxlen = max(len(rows_o), len(rows_b))
    for j in range(min(maxlen, 40)):
        def fmt(rows, j):
            if j >= len(rows): return " "*12, " "*40, ""
            off, style, text, flags = rows[j]
            return style, text, flags

        os_, ot, of = fmt(rows_o, j)
        bs_, bt, bf = fmt(rows_b, j)

        changed = (ot != bt or os_ != bs_)
        marker = "!" if changed else " "
        print(f"{j:4} {os_:12} {repr(ot):40} {marker} {bs_:12} {repr(bt):40}")
        if of or bf:
            print(f"     FLAGS: {of:39} | {bf}")
