"""Dump SEMUA paragraf di sekitar area DAFTAR ISI di file ORIGINAL."""
import zipfile
from lxml import etree

path = r"D:\Freelaces\Test Dafis\Docx 1.docx"
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")
tree = etree.fromstring(xml)
body = tree.find(f"{W}body")
paras = body.findall(f"{W}p")

# Cari DAFTAR ISI
di = None
for i, p in enumerate(paras):
    if "".join(t.text or "" for t in p.iter(f"{W}t")).strip().upper() == "DAFTAR ISI":
        di = i; break

print(f"DAFTAR ISI at [{di}], total paras: {len(paras)}\n")
print("=== Paragraf [di-2 s/d di+10] ORIGINAL ===")
for i in range(max(0, di-2), min(di+11, len(paras))):
    p = paras[i]
    text  = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
    flds  = p.findall(f".//{W}fldChar")
    brs   = p.findall(f".//{W}br")
    pPr   = p.find(f"{W}pPr")
    sn    = pPr.find(f"{W}pStyle") if pPr is not None else None
    style = sn.get(f"{W}val","") if sn is not None else ""
    # Raw XML (truncated)
    raw = etree.tostring(p).decode()
    # Only show if has content
    flags = []
    if flds: flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
    if brs:  flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
    print(f"[{i:3}] style={style:20} | {repr(text[:50]):55} | {' '.join(flags)}")
    if flds or "toc" in raw.lower()[:300]:
        print(f"      XML: {raw[:200]}")