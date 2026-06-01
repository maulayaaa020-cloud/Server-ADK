import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from utils import DocProcessor, is_bab_heading, is_roman_start

path = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
doc  = Document(path)
paras = list(doc.paragraphs)

# Cari posisi KATA PENGANTAR dan para[65]
print("=== KONTEKS SEKITAR para[65] ===")
for i in range(60, 75):
    t = paras[i].text.strip()
    print(f"  para[{i}] style={repr(paras[i].style.name):25s} text={repr(t[:60])}")

print()
print("=== POSISI KATA PENGANTAR ===")
for i, p in enumerate(paras):
    if 'kata pengantar' in p.text.lower():
        print(f"  para[{i}] style={repr(p.style.name):25s} text={repr(p.text.strip()[:60])}")

print()
# Cek apakah para[65] ada SEBELUM atau SESUDAH roman_start_p
proc = DocProcessor(doc, 'Times New Roman', 12)
roman_start_p, bab_p_list = proc.scan_zones()
roman_idx = paras.index(roman_start_p) if roman_start_p in paras else -1
print(f"roman_start_p ada di para[{roman_idx}]: {repr(roman_start_p.text.strip()[:50]) if roman_start_p else None}")
print(f"para[65] ADA {'SEBELUM' if 65 < roman_idx else 'SESUDAH'} roman_start_p")

print()
# Teks para[65] lebih detail
p65 = paras[65]
print(f"=== DETAIL para[65] ===")
print(f"  text  : {repr(p65.text)}")
print(f"  style : {p65.style.name}")
print(f"  runs  : {len(p65.runs)}")
for j, r in enumerate(p65.runs):
    print(f"    run[{j}]: {repr(r.text)}")
# cek apakah ada line break di dalam
from docx.oxml.ns import qn
brs = p65._p.findall('.//' + qn('w:br'))
print(f"  line breaks: {len(brs)}")
for br in brs:
    print(f"    type={br.get(qn('w:type'), 'textWrapping')}")
