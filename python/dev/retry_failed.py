"""Retry Tahap 2 untuk file yang gagal: Docx 8, 9, 10, 11, 12."""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

FAILED = [
    r"D:\Freelaces\Test Dafis\Hasil\Docx 8.docx",
    r"D:\Freelaces\Test Dafis\Hasil\Docx 9.docx",
    r"D:\Freelaces\Test Dafis\Hasil\Docx 10.docx",
    r"D:\Freelaces\Test Dafis\Hasil\Docx 11.docx",
    r"D:\Freelaces\Test Dafis\Hasil\Docx 12.docx",
]

import win32com.client
import pythoncom

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0  # wdAlertsNone — suppress semua dialog

ok = []
err = []

for fpath in FAILED:
    fname = os.path.basename(fpath)
    abs_path = os.path.abspath(fpath)
    try:
        # ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False
        doc = word.Documents.Open(
            abs_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        time.sleep(1)  # beri Word waktu settle

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
            except Exception:
                pass
            try:
                toc_obj = doc.TablesOfContents(i)
                for j in range(1, toc_obj.Range.Paragraphs.Count + 1):
                    p = toc_obj.Range.Paragraphs(j)
                    sname = p.Style.NameLocal.lower().strip()
                    if sname in ("toc 1", "toc1", "daftar isi 1", "toc 11"):
                        p.Range.Font.Bold = True
                    else:
                        p.Range.Font.Bold = False
            except Exception:
                pass

        doc.Save()
        doc.Close(False)
        ok.append(fname)
        print(f"  OK   {fname}")
    except Exception as ex:
        err.append((fname, str(ex)[:100]))
        print(f"  ERR  {fname}  | {ex}")
        try:
            doc.Close(False)
        except Exception:
            pass

word.Quit()
pythoncom.CoUninitialize()

print(f"\nHasil retry: {len(ok)} OK, {len(err)} ERR")
if err:
    print("File yang masih gagal:")
    for f, e in err:
        print(f"  {f}: {e}")
