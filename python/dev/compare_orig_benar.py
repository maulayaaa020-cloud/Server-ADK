"""Bandingkan Original vs File Benar: apa yang berubah di area DAFTAR ISI."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
ORIG = r"D:\Freelaces\Test Dafis"
BENAR = r"D:\Freelaces\Test Dafis\File Benar"

def scan_daftar_area(path, n=20):
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
    for i in range(di, min(di + n, len(paras))):
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
        if instrs: flags.append("instr:" + (instrs[0].text or "")[:30].strip())
        if brs:    flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        if sp is not None:
            t2 = sp.find(f"{W}type")
            flags.append("sect(" + (t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage") + ")")
        result.append((i - di, style, text[:50], flags))
    return di, result

for fname in sorted(os.listdir(ORIG)):
    if not fname.endswith(".docx") or "Docx" not in fname:
        continue
    orig_path = os.path.join(ORIG, fname)
    benar_path = os.path.join(BENAR, fname)
    if not os.path.exists(benar_path):
        print(f"{fname}: File Benar tidak ada")
        continue

    di_o, rows_o = scan_daftar_area(orig_path)
    di_b, rows_b = scan_daftar_area(benar_path)

    print(f"\n{'='*70}")
    print(f"{fname}  |  ORIGINAL di=[{di_o}]  |  FILE BENAR di=[{di_b}]")
    print(f"{'ORIGINAL':<45} | {'FILE BENAR'}")
    print(f"{'-'*45}-+-{'-'*45}")

    maxlen = max(len(rows_o), len(rows_b))
    for j in range(min(maxlen, 12)):
        def fmt(rows, j):
            if j >= len(rows): return " " * 45
            off, style, text, flags = rows[j]
            s = f"+{off} [{style[:12]}] {repr(text[:20]):23} {' '.join(flags)}"
            return s[:45]
        print(f"{fmt(rows_o, j):<45} | {fmt(rows_b, j)}")
