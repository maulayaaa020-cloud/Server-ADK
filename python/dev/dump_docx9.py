import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def show_structure(path, label):
    print(f"\n{'='*60}")
    print(f"  {label}")
    print('='*60)
    doc = Document(path)
    body_els = list(doc.element.body)
    pg = 1
    for i, el in enumerate(body_els):
        tag = el.tag.split('}')[-1]
        if tag != 'p':
            print(f"  [{i:3}] <{tag}>")
            continue
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:50]
        pPr = el.find(f'{{{W}}}pPr')
        has_sect = pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None
        has_pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
        flags = []
        if has_sect: flags.append('SECT')
        if has_pgbr: flags.append('PGBR')
        flag_str = ' [' + ','.join(flags) + ']' if flags else ''
        marker = '  <-- BOUNDARY' if (has_sect or has_pgbr) else ''
        if txt:
            print(f"  [{i:3}] pg{pg}  {repr(txt)}{flag_str}{marker}")
        else:
            print(f"  [{i:3}] pg{pg}  (empty){flag_str}{marker}")
        if has_sect or has_pgbr:
            pg += 1
        if i > 120:
            print("  ... (truncated)")
            break

show_structure(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 9.docx",
    "ORIGINAL"
)
show_structure(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 9_c1.docx",
    "HASIL (after fix)"
)
