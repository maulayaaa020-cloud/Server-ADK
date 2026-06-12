"""Debug Stage 1: cek outlineLvl DAFTAR PUSTAKA dan semua headings yang terdeteksi."""
import sys, subprocess, json, zipfile, tempfile, shutil, os
sys.stdout.reconfigure(encoding='utf-8')
from lxml import etree

PYTHON = r"D:\Freelaces\Server\python.exe"
SCRIPT = r"D:\Freelaces\Server\htdocs\adk\python\daftar_isi.py"
SRC    = r"D:\Freelaces\Test Dafis\Docx 2.docx"

# Tulis ke file temp dulu untuk tidak mengubah output asli
TMPOUT = r"D:\Freelaces\Test Dafis\Hasil\_debug_docx2_s1.docx"

cmd = [PYTHON, SCRIPT, SRC, TMPOUT, "H1+H2+H3", "titik", "Times New Roman", "12", "1.5"]
res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', timeout=120)
try:
    data = json.loads(res.stdout)
    print(f"Status: {data.get('status')}")
    print(f"Heading count: {data.get('heading_count')}")
    print("\nSemua headings:")
    for h in (data.get('headings_preview') or []):
        print(f"  H{h['level']}: {h['text']}")
except Exception:
    print(f"stdout: {res.stdout[:500]}")
    print(f"stderr: {res.stderr[:300]}")

# Cek outlineLvl DAFTAR PUSTAKA di output Stage 1
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

with zipfile.ZipFile(TMPOUT) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(f"{W}body")

print("\n--- DAFTAR PUSTAKA di output Stage 1 ---")
for i, child in enumerate(list(body)):
    txt = "".join((t.text or "") for t in child.findall(f".//{W}t")).strip()
    if "DAFTAR PUSTAKA" not in txt.upper():
        continue
    pPr = child.find(f"{W}pPr")
    pStyle = None
    outlineLvl = None
    if pPr is not None:
        ps = pPr.find(f"{W}pStyle")
        pStyle = ps.get(f"{W}val") if ps is not None else None
        ol = pPr.find(f"{W}outlineLvl")
        outlineLvl = ol.get(f"{W}val") if ol is not None else None
    print(f"  [{i}] style={pStyle!r} | outlineLvl={outlineLvl!r} | {txt[:50]!r}")

# Cleanup
try:
    os.remove(TMPOUT)
except Exception:
    pass
print("\nDone.")
