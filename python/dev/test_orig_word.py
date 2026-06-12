"""
Test: buka ORIG Docx 1, populate SDT dengan TOC field, update via Word.
Tujuan: verifikasi apakah penomoran Roman sudah benar dari ORIG tanpa modifikasi apapun.
"""
import zipfile, os, shutil, win32com.client, pythoncom
from lxml import etree
from lxml import etree as ET

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

ORIG  = r"D:\Freelaces\Test Dafis\Docx 1.docx"
OUT   = r"D:\Freelaces\Test Dafis\Hasil\Docx 1_test_bare.docx"

# Buat copy dari ORIG, lalu populate SDT dengan TOC field (minimal, tanpa modifikasi style)
shutil.copy2(ORIG, OUT)

with zipfile.ZipFile(ORIG) as z:
    xml = z.read("word/document.xml")
root = ET.fromstring(xml)
body = root.find(f"{W}body")
children = list(body)

# Cari SDT
sdt = None
for i, child in enumerate(children):
    if child.tag == f"{W}sdt":
        sdt = child
        print(f"SDT ditemukan di child[{i}]")
        break

if sdt is None:
    print("SDT tidak ditemukan!")
    exit()

# Populate SDT dengan TOC field minimal
sdtContent = sdt.find(f"{W}sdtContent")
if sdtContent is None:
    sdtContent = ET.SubElement(sdt, f"{W}sdtContent")
for child in list(sdtContent):
    sdtContent.remove(child)

# TOCHeading
p_h = ET.SubElement(sdtContent, f"{W}p")
pPr_h = ET.SubElement(p_h, f"{W}pPr")
pStyle_h = ET.SubElement(pPr_h, f"{W}pStyle")
pStyle_h.set(f"{W}val", "TOCHeading")

# TOC field begin + instrText + separate
p1 = ET.SubElement(sdtContent, f"{W}p")
r1a = ET.SubElement(p1, f"{W}r")
fc1 = ET.SubElement(r1a, f"{W}fldChar")
fc1.set(f"{W}fldCharType", "begin")
fc1.set(f"{W}dirty", "true")
r1b = ET.SubElement(p1, f"{W}r")
instr = ET.SubElement(r1b, f"{W}instrText")
instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
instr.text = " TOC \\h \\z \\u "
r1c = ET.SubElement(p1, f"{W}r")
fc2 = ET.SubElement(r1c, f"{W}fldChar")
fc2.set(f"{W}fldCharType", "separate")

# TOC field end
p2 = ET.SubElement(sdtContent, f"{W}p")
r2 = ET.SubElement(p2, f"{W}r")
fc3 = ET.SubElement(r2, f"{W}fldChar")
fc3.set(f"{W}fldCharType", "end")

print("SDT populated dengan TOC field minimal (tanpa modifikasi script).")

# Tulis kembali ke zip
with zipfile.ZipFile(ORIG) as z:
    names = z.namelist()
    infos = {name: z.getinfo(name) for name in names}
    orig_bytes = {name: z.read(name) for name in names}

new_xml = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
import io, zipfile as zf
buf = io.BytesIO()
with zf.ZipFile(buf, 'w', compression=zf.ZIP_DEFLATED) as zout:
    for name in names:
        if name == "word/document.xml":
            zout.writestr(name, new_xml)
        else:
            zout.writestr(infos[name], orig_bytes[name])
with open(OUT, 'wb') as f:
    f.write(buf.getvalue())

print(f"File disimpan: {OUT}")

# Update via Word
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(OUT))
doc.Fields.Update()
for i in range(1, doc.TablesOfContents.Count + 1):
    try: doc.TablesOfContents(i).Update()
    except: pass
doc.Save()
doc.Close(False)
word.Quit()
pythoncom.CoUninitialize()
print("Word update selesai.")

# Scan hasil
with zipfile.ZipFile(OUT) as z:
    xml2 = z.read("word/document.xml")
root2 = ET.fromstring(xml2)
body2 = root2.find(f"{W}body")
children2 = list(body2)
for child in children2:
    if child.tag == f"{W}sdt":
        sdtContent2 = child.find(f"{W}sdtContent")
        inner = sdtContent2.findall(f".//{W}p") if sdtContent2 is not None else []
        print(f"\n=== HASIL BARE (tanpa script): {len(inner)} inner para ===")
        for j, p in enumerate(inner[:8]):
            txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
            pPr = p.find(f"{W}pPr")
            sn  = pPr.find(f"{W}pStyle") if pPr is not None else None
            style = sn.get(f"{W}val","") if sn is not None else ""
            print(f"  [{j}] {style:10} | {repr(txt[:50])}")
        break
