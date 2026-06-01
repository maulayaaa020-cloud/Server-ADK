import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
src = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii'

for n in [2, 4, 5, 7, 10, 16]:
    path = os.path.join(src, f'Docx {n}.docx')
    doc  = Document(path)
    body = list(doc.element.body)
    sects_in_cover = []
    pgbr_in_cover  = []
    roman_start_idx = None
    for i, el in enumerate(body[:60]):
        if not el.tag.endswith('}p'):
            continue
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()
        pPr = el.find(f'{{{W}}}pPr')
        sp  = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
        has_pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
        if sp is not None:
            sects_in_cover.append(i)
        if has_pgbr:
            pgbr_in_cover.append(i)
        # simple roman start detection
        kw = ['kata pengantar', 'daftar isi', 'surat pernyataan', 'lembar persetujuan',
              'abstrak', 'abstract', 'ringkasan']
        if roman_start_idx is None and any(k in txt.lower() for k in kw):
            roman_start_idx = i
            break

    print(f'Docx {n:2d}: cover sectPr at body idx={sects_in_cover} | pgBr at={pgbr_in_cover} | roman kira2 di [{roman_start_idx}]')
