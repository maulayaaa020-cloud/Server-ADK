"""Test: update TOC di File BENAR Docx 1 via Word, apakah page number tetap Roman?"""
import win32com.client, pythoncom, os, zipfile, shutil
from lxml import etree

BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"
COPY  = r"D:\Freelaces\Test Dafis\Hasil\Docx 1_benar_updated.docx"
shutil.copy2(BENAR, COPY)

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

# Update via Word
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(COPY))
doc.Fields.Update()
for i in range(1, doc.TablesOfContents.Count + 1):
    try: doc.TablesOfContents(i).Update()
    except: pass
doc.Save()
doc.Close(False)
word.Quit()
pythoncom.CoUninitialize()
print("Update selesai.")

# Scan
with zipfile.ZipFile(COPY) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(f"{W}body")
for child in list(body):
    if child.tag == f"{W}sdt":
        sdtContent = child.find(f"{W}sdtContent")
        inner = sdtContent.findall(f".//{W}p") if sdtContent else []
        print(f"\n=== BENAR setelah Word update: {len(inner)} inner para ===")
        for j, p in enumerate(inner[:8]):
            txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
            pPr = p.find(f"{W}pPr")
            sn = pPr.find(f"{W}pStyle") if pPr is not None else None
            style = sn.get(f"{W}val","") if sn is not None else ""
            print(f"  [{j}] {style:10} | {repr(txt[:50])}")
        break
