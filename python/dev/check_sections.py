"""Lihat semua sectPr di ORIG dan HASIL Docx 1 untuk debug page numbering."""
import zipfile
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

ORIG  = r"D:\Freelaces\Test Dafis\Docx 1.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"
BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"

def find_all_sections(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    print(f"\n=== {label}: Semua sectPr ===")
    for i, child in enumerate(children):
        # sectPr bisa di body langsung atau di dalam pPr
        if child.tag == f"{W}sectPr":
            print(f"  [BODY] sectPr (final section)")
            pgNum = child.find(f"{W}pgNumType")
            if pgNum is not None:
                fmt = pgNum.get(f"{W}fmt", "decimal")
                start = pgNum.get(f"{W}start", "-")
                print(f"    pgNumType: fmt={fmt}, start={start}")
            pType = child.find(f"{W}type")
            if pType is not None:
                print(f"    type: {pType.get(f'{W}val', 'nextPage')}")
            continue
        if child.tag == f"{W}p":
            pPr = child.find(f"{W}pPr")
            if pPr is not None:
                sp = pPr.find(f"{W}sectPr")
                if sp is not None:
                    txt = "".join(t.text or "" for t in child.iter(f"{W}t")).strip()
                    print(f"  [{i}] para='{txt[:40]}' | sectPr found:")
                    pType = sp.find(f"{W}type")
                    if pType is not None:
                        print(f"    type: {pType.get(f'{W}val', 'nextPage')}")
                    pgNum = sp.find(f"{W}pgNumType")
                    if pgNum is not None:
                        fmt = pgNum.get(f"{W}fmt", "decimal")
                        start = pgNum.get(f"{W}start", "-")
                        print(f"    pgNumType: fmt={fmt}, start={start}")
                    else:
                        print(f"    pgNumType: (tidak ada - lanjut dari section sebelumnya)")
                    # Tampilkan semua child dari sectPr untuk info lengkap
                    for ch in sp:
                        tname = ch.tag.split('}')[1] if '}' in ch.tag else ch.tag
                        if tname not in ('type', 'pgNumType', 'pgSz', 'pgMar'):
                            pass  # skip untuk ringkasan
                    raw = etree.tostring(sp, pretty_print=True).decode()
                    print(f"    raw (first 400): {raw[:400]}")

find_all_sections(ORIG,  "ORIG")
find_all_sections(HASIL, "HASIL")
find_all_sections(BENAR, "BENAR")
