import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RI = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

OUT = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 10_c2.docx'

doc  = Document(OUT)
body = list(doc.element.body)

print('=== sec[0] sectPr XML ===')
print(etree.tostring(doc.sections[0]._sectPr, pretty_print=True).decode())

print('=== sec[1] sectPr XML ===')
print(etree.tostring(doc.sections[1]._sectPr, pretty_print=True).decode())

print('=== Body[16:25] (area sekitar page break) ===')
for i in range(16, min(25, len(body))):
    el = body[i]
    tag = el.tag.split('}')[-1]
    if tag != 'p':
        print(f'  [{i}] <{tag}>')
        continue
    txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:45]
    ppPr = el.find(f'{{{W}}}pPr')
    sp   = ppPr.find(f'{{{W}}}sectPr') if ppPr is not None else None
    flags = []
    if sp is not None:
        t2 = sp.find(f'{{{W}}}type')
        flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
    if any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br')):
        flags.append('pgBr')
    print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')
    if sp is not None:
        print(f'       sectPr: titlePg={sp.find(qn("w:titlePg")) is not None}')
        pn = sp.find(qn('w:pgNumType'))
        if pn is not None:
            print(f'       pgNumType: fmt={pn.get(qn("w:fmt"))} start={pn.get(qn("w:start"))}')
        frefs = [(fr.get(qn('w:type')), fr.get(RI)) for fr in sp.findall(qn('w:footerReference'))]
        hrefs = [(hr.get(qn('w:type')), hr.get(RI)) for hr in sp.findall(qn('w:headerReference'))]
        print(f'       hdr={hrefs}')
        print(f'       ftr={frefs}')
