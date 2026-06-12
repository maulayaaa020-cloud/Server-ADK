"""Investigasi mendalam Docx 1: original vs benar vs hasil."""
import zipfile, os, hashlib
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

ORIG  = r"D:\Freelaces\Test Dafis\Docx 1.docx"
BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"

def md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def filesize(path):
    return os.path.getsize(path)

print("=== FILE SIZE & MD5 ===")
for label, path in [("ORIG ", ORIG), ("BENAR", BENAR), ("HASIL", HASIL)]:
    print(f"  {label}  {filesize(path):,} bytes  md5={md5(path)}")

def para_info(path, start=40, count=20):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")
    paras = body.findall(f"{W}p")
    result = []
    for i in range(start, min(start+count, len(paras))):
        p = paras[i]
        text  = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
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
            flags.append("sect:" + (t2.get(f"{W}val","nextPage") if t2 is not None else "nextPage"))
        result.append((i, style, text, " ".join(flags)))
    return result

print("\n=== AREA DAFTAR ISI (para 46-64) ===")
print(f"{'IDX':>4} {'STYLE':14} {'TEXT':45} {'FLAGS'}")
print("-"*100)
for label, path in [("ORIG ", ORIG), ("BENAR", BENAR), ("HASIL", HASIL)]:
    print(f"\n--- {label} ---")
    for idx, style, text, flags in para_info(path, start=46, count=18):
        mark = ">>>" if text == "DAFTAR ISI" else "   "
        print(f"{mark}{idx:3} {style:14} {repr(text[:43]):47} {flags}")

# Lihat XML raw sekitar DAFTAR ISI di ORIG
print("\n\n=== RAW XML: DAFTAR ISI area di ORIG (para 48-52) ===")
with zipfile.ZipFile(ORIG) as z:
    xml = z.read("word/document.xml")
tree = etree.fromstring(xml)
body = tree.find(f"{W}body")
paras = body.findall(f"{W}p")
for i in range(48, 53):
    p = paras[i]
    txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
    print(f"\n[{i}] text={repr(txt[:50])}")
    print(etree.tostring(p, pretty_print=True).decode()[:800])

print("\n\n=== RAW XML: DAFTAR ISI area di BENAR (para 48-52) ===")
with zipfile.ZipFile(BENAR) as z:
    xml = z.read("word/document.xml")
tree = etree.fromstring(xml)
body = tree.find(f"{W}body")
paras = body.findall(f"{W}p")
for i in range(48, 53):
    p = paras[i]
    txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
    print(f"\n[{i}] text={repr(txt[:50])}")
    print(etree.tostring(p, pretty_print=True).decode()[:800])
