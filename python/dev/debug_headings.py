"""Debug: lihat heading yang terdeteksi dan level-nya, fokus area D.Teknik – E.Sumber."""
import sys, os
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")

from docx import Document
from daftar_isi import detect_headings, get_para_level

ORIG = r"D:\Freelaces\Test Dafis\Docx 1.docx"
doc = Document(ORIG)

headings = detect_headings(doc, 3)

print("=== Semua heading terdeteksi ===")
for idx, lvl, txt in headings:
    marker = " <<<<" if any(x in txt for x in ["Rendah", "Sedang", "Tinggi", "Teknik", "Sumber", "Kendala"]) else ""
    print(f"  [{idx:3}] H{lvl}: {txt[:60]}{marker}")

print(f"\nTotal: {len(headings)} headings")

# Cek paragraf aslinya
print("\n=== Paragraf di sekitar area D.Teknik ===")
for idx, lvl, txt in headings:
    if "Teknik" in txt or "Rendah" in txt or "Sedang" in txt or "Tinggi" in txt or "Sumber" in txt:
        para = doc.paragraphs[idx]
        raw_lvl = get_para_level(para)
        print(f"  [{idx}] H{lvl} (raw={raw_lvl}): '{txt[:60]}' | style: {para.style.name}")
