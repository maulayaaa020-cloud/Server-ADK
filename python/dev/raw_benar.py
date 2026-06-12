"""Lihat raw XML dokumen BENAR Docx 1 - para pertama untuk deteksi penyebab size besar."""
import zipfile
from lxml import etree

BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"
ORIG  = r"D:\Freelaces\Test Dafis\Docx 1.docx"

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

# Lihat namespace apa saja yang ada di BENAR vs ORIG
with zipfile.ZipFile(BENAR) as z:
    xml_b = z.read("word/document.xml")
with zipfile.ZipFile(ORIG) as z:
    xml_o = z.read("word/document.xml")

print(f"ORIG  doc.xml size: {len(xml_o):,} bytes")
print(f"BENAR doc.xml size: {len(xml_b):,} bytes")

# Parse & lihat root element
root_b = etree.fromstring(xml_b)
root_o = etree.fromstring(xml_o)

# Tampilkan namespace yang ada
print(f"\nORIG  nsmap keys: {sorted(root_o.nsmap.keys())[:15]}")
print(f"BENAR nsmap keys: {sorted(root_b.nsmap.keys())[:15]}")

# Lihat paragraf pertama BENAR dalam detail (raw XML, tidak dipotong)
body_b = root_b.find(f"{W}body")
paras_b = body_b.findall(f"{W}p")
print(f"\n=== RAW XML para[0] BENAR (pertama 3000 karakter) ===")
raw0_b = etree.tostring(paras_b[0], pretty_print=True).decode()
print(raw0_b[:3000])
print(f"... total {len(raw0_b)} chars untuk para[0]")

body_o = root_o.find(f"{W}body")
paras_o = body_o.findall(f"{W}p")
print(f"\n=== RAW XML para[0] ORIG (pertama 1000 karakter) ===")
raw0_o = etree.tostring(paras_o[0], pretty_print=True).decode()
print(raw0_o[:1000])
print(f"... total {len(raw0_o)} chars untuk para[0]")

# Cek element unik di BENAR yang tidak ada di ORIG
tags_b = set()
for el in root_b.iter():
    tags_b.add(el.tag.split('}')[1] if '}' in el.tag else el.tag)
tags_o = set()
for el in root_o.iter():
    tags_o.add(el.tag.split('}')[1] if '}' in el.tag else el.tag)

print(f"\n=== Tags di BENAR tapi tidak di ORIG ===")
print(sorted(tags_b - tags_o))
print(f"\n=== Tags di ORIG tapi tidak di BENAR ===")
print(sorted(tags_o - tags_b))

# Hitung jumlah tiap element yang besar
print(f"\n=== Jumlah element terbanyak di BENAR ===")
from collections import Counter
cnt = Counter(el.tag.split('}')[1] if '}' in el.tag else el.tag for el in root_b.iter())
for tag, n in cnt.most_common(20):
    print(f"  {n:8,}  {tag}")
