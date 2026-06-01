import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from utils import DocProcessor, is_bab_heading, is_roman_start, is_false_bab

path = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
doc  = Document(path)
paras = list(doc.paragraphs)

print("=== SCAN SEMUA PARAGRAF — CARI 'BAB' ===\n")
for i, p in enumerate(paras):
    t = p.text.strip()
    if not t:
        continue
    t_low = t.lower()
    # Tampilkan paragraf yang mengandung kata "bab"
    if 'bab' in t_low or 'chapter' in t_low:
        is_bab   = is_bab_heading(t)
        is_false = is_false_bab(p) if is_bab else False
        is_roman = is_roman_start(t)
        print(f"para[{i:3d}]  is_bab={is_bab}  is_false={is_false}  is_roman={is_roman}")
        print(f"         style={repr(p.style.name)}  text={repr(t[:80])}")
        print()

print()
print("=== HASIL scan_zones() ===")
proc = DocProcessor(doc, 'Times New Roman', 12)
roman_start_p, bab_p_list = proc.scan_zones()
print(f"roman_start_p: {repr(roman_start_p.text.strip()[:60]) if roman_start_p else None}")
print(f"bab_p_list ({len(bab_p_list)}):")
for p in bab_p_list:
    print(f"  - {repr(p.text.strip()[:60])}")
