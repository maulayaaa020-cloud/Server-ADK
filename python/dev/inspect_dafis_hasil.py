import zipfile
from lxml import etree

path = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")

tree = etree.fromstring(xml)
body = tree.find(f"{{{ns}}}body")
paras = body.findall(f"{{{ns}}}p")

daftar_idx = None
for i, p in enumerate(paras):
    text = "".join(t.text or "" for t in p.iter(f"{{{ns}}}t")).strip().upper()
    if text == "DAFTAR ISI":
        daftar_idx = i
        break

print(f"DAFTAR ISI at index: {daftar_idx}, total paras: {len(paras)}")
print()

for i in range(daftar_idx, min(daftar_idx + 35, len(paras))):
    p = paras[i]
    text = "".join(t.text or "" for t in p.iter(f"{{{ns}}}t")).strip()
    flds  = p.findall(f".//{{{ns}}}fldChar")
    instrs = p.findall(f".//{{{ns}}}instrText")
    brs   = p.findall(f".//{{{ns}}}br")
    pPr   = p.find(f"{{{ns}}}pPr")
    sectPr = pPr.find(f"{{{ns}}}sectPr") if pPr is not None else None

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

    print(f"[{i:3}] {repr(text[:60]):65} | {' | '.join(flags)}")