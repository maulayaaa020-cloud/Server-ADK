"""Re-run Stage 1 + Stage 2 untuk Docx 2 saja, tanpa menyentuh file lain."""
import os, sys, subprocess, json, time
sys.stdout.reconfigure(encoding='utf-8')

PYTHON  = r"D:\Freelaces\Server\python.exe"
SCRIPT  = r"D:\Freelaces\Server\htdocs\adk\python\daftar_isi.py"
SRC     = r"D:\Freelaces\Test Dafis\Docx 2.docx"
OUT     = r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx"

KEDALAMAN    = "H1+H2+H3"
FORMAT_TITIK = "titik"
FONT         = "Times New Roman"
SIZE_PT      = "12"
SPACING      = "1.5"

print("=== TAHAP 1: Generate TOC (daftar_isi.py) ===")
cmd = [PYTHON, SCRIPT, SRC, OUT, KEDALAMAN, FORMAT_TITIK, FONT, SIZE_PT, SPACING]
res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', timeout=120)
try:
    data = json.loads(res.stdout)
    print(f"  status : {data.get('status')}")
    print(f"  headings: {data.get('heading_count')}")
    print("  Preview:")
    for h in (data.get('headings_preview') or []):
        print(f"    H{h['level']}: {h['text']}")
except Exception:
    print(f"  stdout: {res.stdout[:300]}")
    print(f"  stderr: {res.stderr[:200]}")

if res.returncode != 0:
    print("Tahap 1 gagal. Berhenti.")
    sys.exit(1)

print("\n=== TAHAP 2: Word update + font/bold fix ===")
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

abs_path = os.path.abspath(OUT)
try:
    doc = word.Documents.Open(
        abs_path,
        ConfirmConversions=False,
        ReadOnly=False,
        AddToRecentFiles=False,
    )
    time.sleep(1)
    try:
        doc.Fields.Update()
    except Exception as ex:
        print(f"  Fields.Update skip: {ex}")


    for i in range(1, doc.TablesOfContents.Count + 1):
        try:
            doc.TablesOfContents(i).Update()
        except Exception as ex:
            print(f"  TOC Update warning: {ex}")
        try:
            toc_range = doc.TablesOfContents(i).Range
            toc_range.Font.Name = "Times New Roman"
            toc_range.Font.Size = 12
            print(f"  Font fix applied TOC {i}")
        except Exception as ex:
            print(f"  Font fix warning: {ex}")
        try:
            toc_obj = doc.TablesOfContents(i)
            for j in range(1, toc_obj.Range.Paragraphs.Count + 1):
                p = toc_obj.Range.Paragraphs(j)
                sname = p.Style.NameLocal.lower().strip()
                if sname in ("toc 1", "toc1", "daftar isi 1", "toc 11"):
                    p.Range.Font.Bold = True
                else:
                    p.Range.Font.Bold = False
            print(f"  Bold fix applied TOC {i}")
        except Exception as ex:
            print(f"  Bold fix warning: {ex}")

    doc.Save()
    doc.Close(False)
    print("  OK — file tersimpan.")
except Exception as ex:
    print(f"  ERR: {ex}")
    try:
        doc.Close(False)
    except Exception:
        pass

word.Quit()
pythoncom.CoUninitialize()
print("\n=== SELESAI ===")
