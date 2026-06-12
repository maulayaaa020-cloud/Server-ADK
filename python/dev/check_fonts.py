"""Cek font di Heading char styles dan TOC styles, sebelum dan sesudah script."""
import zipfile, sys, os, shutil
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

def check_styles(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/styles.xml")
    root = etree.fromstring(xml)

    print(f"\n{'='*70}")
    print(f" {label}")
    print(f"{'='*70}")

    for style_el in root.findall(f".//{W}style"):
        stype  = style_el.get(f"{W}type", "")
        sid    = style_el.get(f"{W}styleId", "")
        sname_el = style_el.find(f"{W}name")
        sname  = sname_el.get(f"{W}val", "") if sname_el is not None else ""

        is_toc  = stype == "paragraph" and sname.lower().startswith("toc")
        is_hchar = (stype == "character" and
                    ("Heading" in sname or "Char" in sname or
                     sid.endswith("Char")))

        if not is_toc and not is_hchar:
            continue

        rPr = style_el.find(f"{W}rPr")
        if rPr is None:
            # Check pPr too for paragraph styles
            pPr = style_el.find(f"{W}pPr")
            rPr = pPr.find(f"{W}rPr") if pPr is not None else None

        font_info = "no rPr"
        sz_info   = "?"
        if rPr is not None:
            fonts = rPr.find(f"{W}rFonts")
            sz    = rPr.find(f"{W}sz")
            if fonts is not None:
                ascii_f  = fonts.get(f"{W}ascii", "")
                theme_f  = fonts.get(f"{W}asciiTheme", "")
                font_info = f"ascii={ascii_f!r} theme={theme_f!r}"
            else:
                font_info = "no rFonts"
            sz_info = f"{int(sz.get(f'{W}val','0'))/2:.0f}pt" if sz is not None else "None"

        tag = "TOC " if is_toc else "CHAR"
        print(f"  [{tag}] {sid:20} | {sname:25} | font: {font_info:40} | sz: {sz_info}")


# 1. File Benar asli
check_styles(r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx", "FILE BENAR")

# 2. HASIL setelah script + Word update (current state)
check_styles(r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx", "FILE HASIL (after Word update)")

# 3. Jalankan script pada ORIG dan cek SEBELUM Word update
ORIG = r"D:\Freelaces\Test Dafis\Docx 1.docx"
TEMP = r"D:\Freelaces\Test Dafis\Hasil\_temp_check.docx"

import subprocess
result = subprocess.run(
    [r"D:\Freelaces\Server\python.exe",
     r"D:\Freelaces\Server\htdocs\adk\python\daftar_isi.py",
     ORIG, TEMP, "H1+H2+H3", "titik", "Times New Roman", "12"],
    capture_output=True, text=True
)
print(f"\nScript output: {result.stdout.strip()}")
if result.returncode != 0:
    print(f"STDERR: {result.stderr[:300]}")

check_styles(TEMP, "TEMP (setelah script, SEBELUM Word update)")

# Juga cek rStyle yang dipakai di TOC entries dalam TEMP
with zipfile.ZipFile(TEMP) as z:
    doc_xml = z.read("word/document.xml")
root_doc = etree.fromstring(doc_xml)
body = root_doc.find(f"{W}body")
sdt = None
for child in list(body):
    if child.tag == f"{W}sdt":
        sdt = child
        break
if sdt:
    sdtContent = sdt.find(f"{W}sdtContent")
    paras = sdtContent.findall(f".//{W}p") if sdtContent else []
    print(f"\n  TEMP TOC runs rStyle/font (first 5 entries):")
    for p in paras[1:6]:
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()[:30]
        for r in p.findall(f"{W}r"):
            rPr = r.find(f"{W}rPr")
            if rPr is None: continue
            rStyle = rPr.find(f"{W}rStyle")
            fonts  = rPr.find(f"{W}rFonts")
            rs = rStyle.get(f"{W}val") if rStyle is not None else "none"
            fa = fonts.get(f"{W}ascii","") if fonts is not None else ""
            ft = fonts.get(f"{W}asciiTheme","") if fonts is not None else ""
            print(f"    '{txt}' | rStyle={rs} | ascii={fa!r} theme={ft!r}")
            break

os.remove(TEMP)
