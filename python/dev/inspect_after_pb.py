"""
Cek semua file Hasil: setelah page break pertama sesudah DAFTAR ISI,
adakah paragraf yang masih punya konten mencurigakan (fldChar, teks, dll)?
"""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

hasil_dir = r"D:\Freelaces\Test Dafis\Hasil"

def tag(el): return el.tag.split('}')[-1] if '}' in el.tag else el.tag

def para_info(p):
    text   = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
    flds   = p.findall(f".//{W}fldChar")
    instrs = p.findall(f".//{W}instrText")
    brs    = p.findall(f".//{W}br")
    pPr    = p.find(f"{W}pPr")
    sectPr = pPr.find(f"{W}sectPr") if pPr is not None else None
    flags  = []
    if flds:
        flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
    if instrs:
        flags.append("instr:" + (instrs[0].text or "")[:30].strip())
    if brs:
        flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
    if sectPr is not None:
        t2 = sectPr.find(f"{W}type")
        flags.append("sectPr(" + (t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage") + ")")
    return text, flags

for fname in sorted(os.listdir(hasil_dir)):
    if not fname.endswith(".docx"):
        continue
    path = os.path.join(hasil_dir, fname)
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")
    paras = body.findall(f"{W}p")

    # Cari DAFTAR ISI
    daftar_idx = None
    for i, p in enumerate(paras):
        if "".join(t.text or "" for t in p.iter(f"{W}t")).strip().upper() == "DAFTAR ISI":
            daftar_idx = i
            break
    if daftar_idx is None:
        continue

    # Cari page break pertama setelah DAFTAR ISI
    pb_idx = None
    for i in range(daftar_idx + 1, min(daftar_idx + 15, len(paras))):
        brs = paras[i].findall(f".//{W}br")
        if any(b.get(f"{W}type") == "page" for b in brs):
            pb_idx = i
            break
        pPr = paras[i].find(f"{W}pPr")
        if pPr is not None:
            sp = pPr.find(f"{W}sectPr")
            if sp is not None:
                t2 = sp.find(f"{W}type")
                val = t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage"
                if val != "continuous":
                    pb_idx = i
                    break

    if pb_idx is None:
        print(f"{fname}: TIDAK ADA page break setelah DAFTAR ISI")
        continue

    # Dump 5 paragraf setelah page break
    problems = []
    for i in range(pb_idx + 1, min(pb_idx + 6, len(paras))):
        text, flags = para_info(paras[i])
        if text or flags:
            problems.append(f"  [{i}] {repr(text[:50]):55} | {' | '.join(flags)}")

    if problems:
        print(f"\n{fname} (DAFTAR ISI={daftar_idx}, pb={pb_idx}):")
        for line in problems:
            print(line)
    else:
        print(f"{fname}: OK (bersih setelah page break)")
