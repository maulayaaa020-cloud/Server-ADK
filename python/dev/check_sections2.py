"""Lihat full sectPr XML untuk memahami page numbering secara detail."""
import zipfile
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

ORIG  = r"D:\Freelaces\Test Dafis\Docx 1.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"
BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"

def show_sections(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    print(f"\n{'='*60}")
    print(f"{label}")

    sects = []
    for i, child in enumerate(children):
        if child.tag == f"{W}sectPr":
            sects.append((i, "(body)", child))
        elif child.tag == f"{W}p":
            pPr = child.find(f"{W}pPr")
            if pPr is not None:
                sp = pPr.find(f"{W}sectPr")
                if sp is not None:
                    txt = "".join(t.text or "" for t in child.iter(f"{W}t")).strip()
                    sects.append((i, f"para='{txt[:25]}'", sp))

    for idx, desc, sp in sects:
        pgNum = sp.find(f"{W}pgNumType")
        pType = sp.find(f"{W}type")
        fmt_  = pgNum.get(f"{W}fmt", "decimal") if pgNum is not None else "(inherit)"
        start_ = pgNum.get(f"{W}start", "-") if pgNum is not None else "-"
        type_ = pType.get(f"{W}val", "nextPage") if pType is not None else "(none)"
        chapStart = sp.find(f"{W}chapNumLvl")
        print(f"  [{idx:3}] {desc:30} | type={type_:10} fmt={fmt_:12} start={start_}")

# Lihat juga: untuk section 2 (paras 49..115) di ORIG, berapa halaman lowerRoman?
# Kita perlu lihat apakah section 2 punya pgNumType sendiri
def show_full_sect_at(path, label, idx):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)
    child = children[idx]
    pPr = child.find(f"{W}pPr")
    if pPr is not None:
        sp = pPr.find(f"{W}sectPr")
        if sp is not None:
            raw = etree.tostring(sp, pretty_print=True).decode()
            # Hapus namespace declarations yang panjang untuk keterbacaan
            import re
            raw_clean = re.sub(r'\s+xmlns:[a-z0-9]+="[^"]*"', '', raw)
            print(f"\n=== {label} sectPr at [{idx}] (cleaned) ===")
            print(raw_clean[:1500])

show_sections(ORIG,  "ORIG")
show_sections(HASIL, "HASIL")
show_sections(BENAR, "BENAR")

# Lihat sectPr di para[115] ORIG dan BENAR (section boundary setelah DAFTAR ISI)
show_full_sect_at(ORIG,  "ORIG",  115)
show_full_sect_at(BENAR, "BENAR", 115)
