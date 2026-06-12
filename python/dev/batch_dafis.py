"""
batch_dafis.py — Batch process 20 file:
1. Jalankan daftar_isi.py pada tiap file original
2. Buka output di Word via win32com, update fields, save

Usage: python batch_dafis.py
"""
import os, sys, subprocess, time

sys.stdout.reconfigure(encoding='utf-8')

PYTHON     = r"D:\Freelaces\Server\python.exe"
SCRIPT     = r"D:\Freelaces\Server\htdocs\adk\python\daftar_isi.py"
SRC_DIR    = r"D:\Freelaces\Test Dafis"
OUT_DIR    = r"D:\Freelaces\Test Dafis\Hasil"

# Parameter daftar_isi.py
KEDALAMAN    = "H1+H2+H3"
FORMAT_TITIK = "titik"
FONT         = "Times New Roman"
SIZE_PT      = "12"
SPACING      = "1.5"

os.makedirs(OUT_DIR, exist_ok=True)

files = sorted(
    f for f in os.listdir(SRC_DIR)
    if f.lower().endswith(".docx") and not f.startswith("~") and f.startswith("Docx")
)

print(f"Ditemukan {len(files)} file\n")

ok_count = err_count = 0
processed = []

# ── Tahap 1: Jalankan daftar_isi.py ──────────────────────────────────────────
print("=== TAHAP 1: Generate TOC field dengan daftar_isi.py ===")
for fname in files:
    inp = os.path.join(SRC_DIR, fname)
    out = os.path.join(OUT_DIR, fname)
    cmd = [PYTHON, SCRIPT, inp, out, KEDALAMAN, FORMAT_TITIK, FONT, SIZE_PT, SPACING]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', timeout=120)
        stdout = (res.stdout or "").strip()
        stderr = (res.stderr or "").strip()
        # Coba parse JSON
        import json as _json
        try:
            data = _json.loads(stdout)
            status = data.get("status", "unknown")
        except Exception:
            status = "ok" if res.returncode == 0 else "error"
            data = {"message": (stderr or stdout)[:200]}

        if status in ("success", "ok") or res.returncode == 0:
            ok_count += 1
            processed.append(out)
            print(f"  OK   {fname}")
        else:
            err_count += 1
            msg = data.get("message", stderr or stdout)[:100]
            print(f"  ERR  {fname}  | {msg}")
    except subprocess.TimeoutExpired:
        err_count += 1
        print(f"  TIMEOUT  {fname}")
    except Exception as ex:
        err_count += 1
        print(f"  EX  {fname}  | {ex}")

print(f"\nTahap 1 selesai: {ok_count} OK, {err_count} ERR\n")

if not processed:
    print("Tidak ada file yang berhasil di-proses. Berhenti.")
    sys.exit(1)

# ── Tahap 2: Update fields via win32com ───────────────────────────────────────
print("=== TAHAP 2: Update TOC fields di Word ===")
try:
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    for fpath in processed:
        fname = os.path.basename(fpath)
        abs_path = os.path.abspath(fpath)
        try:
            doc = word.Documents.Open(abs_path)
            # Update semua fields (termasuk TOC)
            doc.Fields.Update()
            # Update TOC spesifik + fix font setelah update
            for i in range(1, doc.TablesOfContents.Count + 1):
                try:
                    doc.TablesOfContents(i).Update()
                except Exception:
                    pass
                # Word menyalin theme font dari heading runs ke TOC entries.
                # Override eksplisit ke Times New Roman 12pt setelah update.
                try:
                    toc_range = doc.TablesOfContents(i).Range
                    toc_range.Font.Name = "Times New Roman"
                    toc_range.Font.Size = 12
                except Exception:
                    pass
                # Fix bold: TOC1 (BAB) harus bold, TOC2/3 tidak bold.
                # Word kadang menaruh bold=False eksplisit dari heading char style.
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
            print(f"  OK   {fname}")
        except Exception as ex:
            print(f"  ERR  {fname}  | {ex}")
            try:
                doc.Close(False)
            except Exception:
                pass

    word.Quit()
    pythoncom.CoUninitialize()
    print("\nTahap 2 selesai.")

except Exception as ex:
    print(f"win32com gagal: {ex}")
    print("File tetap tersimpan dari tahap 1 (tanpa Word update).")

print(f"\n=== SELESAI ===")
print(f"Output tersimpan di: {OUT_DIR}")
