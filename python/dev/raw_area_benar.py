"""Lihat raw XML sekitar DAFTAR ISI di BENAR Docx 1 — untuk memahami struktur TOC."""
import zipfile
from lxml import etree

BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

with zipfile.ZipFile(BENAR) as z:
    xml_b = z.read("word/document.xml")

root = etree.fromstring(xml_b)
body = root.find(f"{W}body")

# Cari semua direct children dari body (termasuk tbl, sdt, dll)
children = list(body)
print(f"Total direct children body: {len(children)}")
print("Tipe children:")
from collections import Counter
types = Counter(c.tag.split('}')[1] if '}' in c.tag else c.tag for c in children)
print(types.most_common(10))

# Cari posisi DAFTAR ISI
di_idx = None
for i, child in enumerate(children):
    if child.tag == f"{W}p":
        txt = "".join(t.text or "" for t in child.iter(f"{W}t")).strip().upper()
        if txt == "DAFTAR ISI":
            di_idx = i
            print(f"\nDAFTAR ISI found at child index {i} (para)")
            break

if di_idx is not None:
    print(f"\n=== children[{di_idx-2} .. {di_idx+8}] ===")
    for i in range(max(0, di_idx-2), min(di_idx+8, len(children))):
        child = children[i]
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
        if tag == "p":
            txt = "".join(t.text or "" for t in child.iter(f"{W}t")).strip()
            flds = child.findall(f".//{W}fldChar")
            instrs = child.findall(f".//{W}instrText")
            wh = child.findall(f".//{W}webHidden")
            bm = child.findall(f".//{W}bookmarkStart")
            flags = []
            if flds:   flags.append(f"fld:{len(flds)}")
            if instrs: flags.append(f"instr:{(instrs[0].text or '')[:20]}")
            if wh:     flags.append(f"webHidden:{len(wh)}")
            if bm:     flags.append(f"bm:{[b.get(f'{W}name','') for b in bm]}")
            print(f"  [{i}] <p>  text={repr(txt[:50]):55} {' '.join(flags)}")
        else:
            # Untuk non-p, tampilkan raw singkat
            raw = etree.tostring(child, pretty_print=True).decode()
            print(f"  [{i}] <{tag}>  ({len(raw)} chars)")
            print(f"       {raw[:200]}")

# Cari semua fldChar di dalam body dan lokasi paragrafnya
print(f"\n=== Lokasi fldChar di dokumen ===")
para_with_fld = []
for i, child in enumerate(children):
    if child.tag == f"{W}p":
        flds = child.findall(f".//{W}fldChar")
        if flds:
            txt = "".join(t.text or "" for t in child.iter(f"{W}t")).strip()
            instrs = child.findall(f".//{W}instrText")
            types_fld = [f.get(f"{W}fldCharType","?") for f in flds]
            instr_txt = (instrs[0].text or "")[:30].strip() if instrs else ""
            para_with_fld.append((i, types_fld, txt[:40], instr_txt))

print(f"Total paras dengan fldChar: {len(para_with_fld)}")
for idx, types_f, txt, instr in para_with_fld[:20]:
    print(f"  [{idx}] fld={types_f} | {repr(txt):45} | instr={repr(instr)}")

# Lihat raw XML dari para DAFTAR ISI (para[49]) di BENAR
print(f"\n\n=== RAW XML para[49] (DAFTAR ISI) di BENAR ===")
raw49 = etree.tostring(children[di_idx], pretty_print=True).decode()
print(raw49[:5000])
print(f"... total {len(raw49)} chars")
