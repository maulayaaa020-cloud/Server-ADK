import zipfile
from lxml import etree

# Cek file original dan cari TOC field di mana saja
for fname in ["Docx 1.docx"]:
    path = rf"D:\Freelaces\Test Dafis\{fname}"
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")

    tree = etree.fromstring(xml)
    body = tree.find(f"{{{ns}}}body")
    paras = body.findall(f"{{{ns}}}p")

    print(f"=== {fname} | total paras: {len(paras)} ===")
    for i, p in enumerate(paras):
        text  = "".join(t.text or "" for t in p.iter(f"{{{ns}}}t")).strip()
        flds  = p.findall(f".//{{{ns}}}fldChar")
        instrs = p.findall(f".//{{{ns}}}instrText")
        brs   = p.findall(f".//{{{ns}}}br")
        pPr   = p.find(f"{{{ns}}}pPr")
        sectPr = pPr.find(f"{{{ns}}}sectPr") if pPr is not None else None

        # Hanya print baris yang relevan
        is_interesting = (
            flds or instrs or brs or sectPr is not None
            or text.upper() in ("DAFTAR ISI", "KATA PENGANTAR")
            or text.startswith("BAB ")
        )
        if not is_interesting:
            continue

        flags = []
        if flds:
            flags.append("fld:" + ",".join(f.get(f"{{{ns}}}fldCharType", "?") for f in flds))
        if instrs:
            flags.append("instr:" + (instrs[0].text or "")[:40].strip())
        if brs:
            flags.append("br:" + ",".join(b.get(f"{{{ns}}}type", "soft") for b in brs))
        if sectPr is not None:
            t2 = sectPr.find(f"{{{ns}}}type")
            sv = t2.get(f"{{{ns}}}val", "nextPage") if t2 is not None else "nextPage"
            flags.append(f"sectPr({sv})")

        print(f"  [{i:3}] {repr(text[:60]):65} | {' | '.join(flags)}")