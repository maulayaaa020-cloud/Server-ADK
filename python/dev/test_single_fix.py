"""Test fix font + alpha context pada Docx 1 saja."""
import os, sys, subprocess, json
sys.stdout.reconfigure(encoding='utf-8')

PYTHON  = r"D:\Freelaces\Server\python.exe"
SCRIPT  = r"D:\Freelaces\Server\htdocs\adk\python\daftar_isi.py"
ORIG    = r"D:\Freelaces\Test Dafis\Docx 1.docx"
OUT     = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"

# Tahap 1
cmd = [PYTHON, SCRIPT, ORIG, OUT, "H1+H2+H3", "titik", "Times New Roman", "12", "1.5"]
res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
try:
    data = json.loads(res.stdout)
    print(f"Tahap 1: {data['status']} | {data['heading_count']} headings")
    print("Headings preview:")
    for h in data['headings_preview']:
        print(f"  H{h['level']}: {h['text']}")
except Exception:
    print(f"Tahap 1 output: {res.stdout[:300]}")
    print(f"STDERR: {res.stderr[:200]}")

# Tahap 2: Word update + font fix
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(OUT))
doc.Fields.Update()
for i in range(1, doc.TablesOfContents.Count + 1):
    try:
        doc.TablesOfContents(i).Update()
    except Exception:
        pass
    try:
        toc_range = doc.TablesOfContents(i).Range
        toc_range.Font.Name = "Times New Roman"
        toc_range.Font.Size = 12
        print(f"Font fix applied pada TOC {i}")
    except Exception as ex:
        print(f"Font fix gagal: {ex}")
    try:
        toc_obj = doc.TablesOfContents(i)
        for j in range(1, toc_obj.Range.Paragraphs.Count + 1):
            p = toc_obj.Range.Paragraphs(j)
            sname = p.Style.NameLocal.lower().strip()
            if sname in ("toc 1", "toc1", "daftar isi 1", "toc 11"):
                p.Range.Font.Bold = True
            else:
                p.Range.Font.Bold = False
        print(f"Bold fix applied pada TOC {i}")
    except Exception as ex:
        print(f"Bold fix gagal: {ex}")
doc.Save()
doc.Close(False)
word.Quit()
pythoncom.CoUninitialize()
print("Tahap 2 selesai.")

# Verifikasi — cek TOC entries
import zipfile
from lxml import etree
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

with zipfile.ZipFile(OUT) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(f"{W}body")
sdt = None
for child in list(body):
    if child.tag == f"{W}sdt":
        sdt = child
        break

if sdt:
    sdtContent = sdt.find(f"{W}sdtContent")
    paras = sdtContent.findall(f".//{W}p") if sdtContent else []
    print(f"\n=== Hasil ({len(paras)} paragraf di TOC) ===")
    for p in paras[1:]:  # skip TOCHeading
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if not txt:
            continue
        pPr = p.find(f"{W}pPr")
        pStyle = pPr.find(f"{W}pStyle") if pPr else None
        style = pStyle.get(f"{W}val","") if pStyle is not None else ""
        # Cek font run pertama
        font_info = "?"
        for r in p.findall(f"{W}r"):
            rPr = r.find(f"{W}rPr")
            if rPr is None: continue
            fonts = rPr.find(f"{W}rFonts")
            sz    = rPr.find(f"{W}sz")
            if fonts is not None:
                fa = fonts.get(f"{W}ascii","")
                ft = fonts.get(f"{W}asciiTheme","")
                font_info = f"ascii={fa!r}" if fa else f"theme={ft!r}"
            sz_pt = f"{int(sz.get(f'{W}val','0'))/2:.0f}pt" if sz is not None else "?"
            print(f"  [{style:5}] font={font_info:25} sz={sz_pt:5} | {txt[:50]}")
            break
