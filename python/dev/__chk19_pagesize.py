import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn

def emu_to_cm(emu):
    return round(emu / 914400 * 2.54, 2)

def check_pagesize(label, path):
    doc = Document(path)
    print(f"\n=== {label} ===")
    for i, sec in enumerate(doc.sections):
        pgSz = sec._sectPr.find(qn('w:pgSz'))
        if pgSz is None:
            print(f"  Section {i}: pgSz tidak ada")
            continue
        w     = int(pgSz.get(qn('w:w'), 0))
        h     = int(pgSz.get(qn('w:h'), 0))
        orient = pgSz.get(qn('w:orient'), 'portrait')
        w_cm  = emu_to_cm(w * 635)   # twips → EMU
        h_cm  = emu_to_cm(h * 635)
        # Identifikasi ukuran kertas
        if 2050 <= w <= 2100 and 2960 <= h <= 3000:
            paper = "A4"
        elif 2100 <= w <= 2160 and 2780 <= h <= 2810:
            paper = "Letter"
        elif 2480 <= w <= 2490 and 3500 <= h <= 3510:
            paper = "A3"
        else:
            paper = "custom"
        print(f"  Section {i}: {w}x{h} twips ({w_cm}x{h_cm} cm) orient={orient} = {paper}")

base = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii'
check_pagesize("ASLI", f"{base}\\Docx 19.docx")
check_pagesize("HASIL", f"{base}\\hasil\\Docx 19.docx")
