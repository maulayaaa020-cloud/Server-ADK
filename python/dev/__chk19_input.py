import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

path = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
doc = Document(path)
paras = list(doc.paragraphs)

print(f"Total paragraf: {len(paras)}")
print()

# Cari semua section break dan tampilkan konteksnya
print("=== SECTION BREAKS DI DOKUMEN ASLI ===")
breaks = [i for i, p in enumerate(paras) if DocProcessor._has_sectPr(p._p)]
print(f"Jumlah section break: {len(breaks)}")
print()

for bi, b in enumerate(breaks):
    # Ambil 3 paragraf sebelum dan sesudah break
    before = []
    for i in range(max(0, b-3), b+1):
        t = paras[i].text.strip()
        flag = " <-- BREAK" if i == b else ""
        before.append(f"  [{i}] {repr(t[:60])}{flag}")
    after = []
    for i in range(b+1, min(len(paras), b+4)):
        t = paras[i].text.strip()
        after.append(f"  [{i}] {repr(t[:60])}")
    print(f"--- Break #{bi+1} (para index {b}) ---")
    for line in before:
        print(line)
    print("  ---")
    for line in after:
        print(line)
    print()
