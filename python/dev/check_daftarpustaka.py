"""Cek outline level dan semua properties paragraf DAFTAR PUSTAKA di Docx 2 original dan Hasil."""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import zipfile
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
SEP = "=" * 70

def check_para_detail(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    print(f"\n{SEP}")
    print(f" {label}")
    print(SEP)
    print("\n--- DAFTAR PUSTAKA ---")

    for i, child in enumerate(children):
        txt = "".join((t.text or "") for t in child.findall(f".//{W}t")).strip()
        if "DAFTAR PUSTAKA" not in txt.upper():
            continue

        pPr = child.find(f"{W}pPr")
        if pPr is None:
            print(f"  [{i}] pPr=None | txt={txt[:50]!r}")
            continue

        # pStyle
        pStyle = pPr.find(f"{W}pStyle")
        style_val = pStyle.get(f"{W}val") if pStyle is not None else None

        # outlineLvl
        outlineLvl = pPr.find(f"{W}outlineLvl")
        outline_val = outlineLvl.get(f"{W}val") if outlineLvl is not None else None

        # numPr
        numPr = pPr.find(f"{W}numPr")
        num_info = None
        if numPr is not None:
            ilvl = numPr.find(f"{W}ilvl")
            numId = numPr.find(f"{W}numId")
            num_info = f"ilvl={ilvl.get(f'{W}val') if ilvl is not None else '?'}, numId={numId.get(f'{W}val') if numId is not None else '?'}"

        # Dump full pPr XML
        ppr_xml = etree.tostring(pPr, pretty_print=True).decode()

        print(f"\n  [{i}] style={style_val!r} | outlineLvl={outline_val!r} | numPr={num_info}")
        print(f"  txt: {txt[:60]!r}")
        print(f"  pPr XML (first 800 chars):")
        for line in ppr_xml[:800].split("\n"):
            print(f"    {line}")


check_para_detail(r"D:\Freelaces\Test Dafis\Docx 2.docx", "ORIGINAL Docx 2")
check_para_detail(r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx", "HASIL Docx 2")
check_para_detail(r"D:\Freelaces\Test Dafis\File Benar\Docx 2.docx", "FILE BENAR Docx 2")
