"""Update TOC via Word untuk satu file, lalu scan hasilnya."""
import win32com.client, pythoncom, zipfile, os
from lxml import etree

FILE = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"
BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"

# Update via Word
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(FILE))
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
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

def scan_sdt(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    for i, child in enumerate(children):
        if child.tag == f"{W}sdt":
            sdtContent = child.find(f"{W}sdtContent")
            inner = sdtContent.findall(f".//{W}p") if sdtContent else []
            print(f"\n=== {label}: SDT di child[{i}], {len(inner)} inner para ===")
            for j, p in enumerate(inner[:10]):
                txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
                flds = p.findall(f".//{W}fldChar")
                instrs = p.findall(f".//{W}instrText")
                pPr = p.find(f"{W}pPr")
                sn  = pPr.find(f"{W}pStyle") if pPr is not None else None
                style = sn.get(f"{W}val","") if sn is not None else ""
                flags = ""
                if flds: flags += "fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds)
                if instrs: flags += " instr:" + (instrs[0].text or "")[:25]
                print(f"  [{j}] {style:12} | {repr(txt[:45]):50} | {flags}")
            if len(inner) > 10:
                print(f"  ... ({len(inner)-10} lagi)")
            break

scan_sdt(FILE,  "HASIL (updated)")
scan_sdt(BENAR, "BENAR")
