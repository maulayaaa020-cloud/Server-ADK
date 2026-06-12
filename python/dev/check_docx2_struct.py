"""Cek struktur body Docx 2 Hasil - cari SDT dan TOC field."""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import zipfile
from lxml import etree

path = r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx"
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{ns}}}"

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(f"{W}body")

children = list(body)
print(f"Total anak body: {len(children)}")

# Cari SDT
sdts = body.findall(f".//{W}sdt")
print(f"Total SDT (nested): {len(sdts)}")
for i, sdt in enumerate(sdts[:5]):
    txts = sdt.findall(f".//{W}t")
    txt = " ".join((t.text or "") for t in txts[:5])[:60]
    print(f"  SDT [{i}]: {txt!r}")

# Cari TOC field
instrs = body.findall(f".//{W}instrText")
print(f"Total instrText: {len(instrs)}")
for instr in instrs[:5]:
    t = (instr.text or "")[:60]
    print(f"  instr: {t!r}")

# Cari paragraf DAFTAR ISI
print("\nCari paragraf 'DAFTAR ISI':")
for i, child in enumerate(children):
    txts = child.findall(f".//{W}t")
    txt = "".join((t.text or "") for t in txts)
    if "DAFTAR ISI" in txt.upper():
        tag = child.tag.split("}")[-1]
        print(f"  [{i}] {tag}: {txt[:60]!r}")
