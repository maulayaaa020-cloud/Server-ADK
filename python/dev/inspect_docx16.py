import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

SRC = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 16.docx'

doc = Document(SRC)
body = doc.element.body

print('=== Raw body children (first 60) ===')
for i, el in enumerate(body):
    if i >= 60: break
    tag = el.tag.split('}')[-1]
    if tag == 'sectPr':
        print(f'  [{i}] <FLOATING sectPr>')
        continue
    if tag == 'tbl':
        print(f'  [{i}] <tbl>')
        continue
    if tag != 'p':
        print(f'  [{i}] <{tag}>')
        continue
    # paragraph
    txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:35]
    pPr = el.find(f'{{{W}}}pPr')
    sp = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
    has_pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
    flags = []
    if sp is not None:
        t2 = sp.find(f'{{{W}}}type')
        flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
    if has_pgbr:
        flags.append('pgBr')
    print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')

print('\n=== Total body children:', len(list(body)))
