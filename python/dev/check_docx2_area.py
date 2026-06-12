"""Lihat paragraf sekitar DAFTAR ISI di Docx 2 Hasil."""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import zipfile
from lxml import etree

path = r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx"
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(f"{W}body")
children = list(body)

print(f"Total anak body: {len(children)}")
print("\nParagraf 140-175 (area DAFTAR ISI):")
for i in range(140, min(175, len(children))):
    child = children[i]
    tag = child.tag.split("}")[-1]
    txt = "".join((t.text or "") for t in child.findall(f".//{W}t")).strip()
    pPr = child.find(f"{W}pPr")
    pStyle = None
    if pPr is not None:
        ps = pPr.find(f"{W}pStyle")
        if ps is not None:
            pStyle = ps.get(f"{W}val")
    instrs = child.findall(f".//{W}instrText")
    instr = (instrs[0].text or "")[:30] if instrs else ""
    print(f"  [{i}] {tag:4} style={str(pStyle):15} | instr={instr!r:30} | {txt[:50]!r}")
