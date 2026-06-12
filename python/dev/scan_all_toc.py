"""Scan semua file output untuk menemukan SEMUA fldChar/instrText TOC di mana saja."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

for fname in ["Docx 1.docx"]:
    for folder, label in [
        (r"D:\Freelaces\Test Dafis", "ORIGINAL"),
        (r"D:\Freelaces\Test Dafis\Hasil", "OUTPUT"),
    ]:
        path = os.path.join(folder, fname)
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml")
        tree = etree.fromstring(xml)
        body = tree.find(f"{W}body")
        paras = body.findall(f"{W}p")

        print(f"\n=== {label}: {fname} ({len(paras)} paras) ===")
        for i, p in enumerate(paras):
            flds   = p.findall(f".//{W}fldChar")
            instrs = p.findall(f".//{W}instrText")
            if not flds and not instrs:
                continue
            text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
            fld_types = ",".join(f.get(f"{W}fldCharType","?") for f in flds)
            instr_txt = " | ".join((i.text or "")[:50].strip() for i in instrs)
            print(f"  [{i:3}] fld:{fld_types:25} | instr: {instr_txt[:60]} | text: {repr(text[:40])}")