import zipfile
from lxml import etree

path = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")

tree = etree.fromstring(xml)
body = tree.find(f"{{{ns}}}body")
paras = body.findall(f"{{{ns}}}p")

# Cari DAFTAR ISI dan dump paragraf 45-65 lengkap
daftar_idx = None
for i, p in enumerate(paras):
    text = "".join(t.text or "" for t in p.iter(f"{{{ns}}}t")).strip().upper()
    if text == "DAFTAR ISI":
        daftar_idx = i
        break

print(f"DAFTAR ISI at [{daftar_idx}]\n")

# Dump full XML paragraf 49 sd 55
for i in range(daftar_idx, min(daftar_idx + 8, len(paras))):
    p = paras[i]
    raw = etree.tostring(p, pretty_print=True).decode()
    print(f"=== [{i}] ===")
    print(raw[:800])
    print()
