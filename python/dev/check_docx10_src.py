import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

SRC = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\Docx 10.docx'

doc = Document(SRC)
body = list(doc.element.body)

print('=== Docx 10 source body [0:60] ===')
for i in range(min(60, len(body))):
    el = body[i]
    tag = el.tag.split('}')[-1]
    if tag == 'tbl':
        print(f'  [{i}] <tbl>')
        continue
    if tag != 'p':
        print(f'  [{i}] <{tag}>')
        continue
    txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:45]
    ppPr = el.find(f'{{{W}}}pPr')
    sp2 = ppPr.find(f'{{{W}}}sectPr') if ppPr is not None else None
    flags = []
    if sp2 is not None:
        t2 = sp2.find(f'{{{W}}}type')
        flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
    if any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br')):
        flags.append('pgBr')
    print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')
