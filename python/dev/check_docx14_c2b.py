import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

OUT = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 14_c2.docx'
doc  = Document(OUT)
body = list(doc.element.body)

# Cek footer XML sec[2] (roman zone)
print('=== Footer XML sec[2] (roman zone) ===')
sec2 = doc.sections[2]
print(etree.tostring(sec2.footer._element, pretty_print=True).decode()[:1500])

# Body lanjutan [50:90]
print('\n=== Body [50:90] ===')
for i in range(50, min(90, len(body))):
    el  = body[i]
    tag = el.tag.split('}')[-1]
    if tag != 'p':
        print(f'  [{i}] <{tag}>')
        continue
    txt   = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:45]
    ppPr  = el.find(f'{{{W}}}pPr')
    sp    = ppPr.find(f'{{{W}}}sectPr') if ppPr is not None else None
    flags = []
    if sp is not None:
        t2 = sp.find(f'{{{W}}}type')
        flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
        pn2 = sp.find(qn('w:pgNumType'))
        if pn2 is not None:
            flags.append(f'pgNum(fmt={pn2.get(qn("w:fmt"))} start={pn2.get(qn("w:start"))})')
    if any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br')):
        flags.append('pgBr')
    print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')
