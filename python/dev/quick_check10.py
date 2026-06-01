import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
doc = Document(r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 10_c2.docx')
body = list(doc.element.body)
print('Total sections:', len(list(doc.sections)))
for i, sec in enumerate(doc.sections):
    pn = sec._sectPr.find(qn('w:pgNumType'))
    fmt   = pn.get(qn('w:fmt'))   if pn is not None else None
    start = pn.get(qn('w:start')) if pn is not None else None
    ftr   = len(sec._sectPr.findall(qn('w:footerReference')))
    print(f'  sec[{i}] fmt={fmt} start={start} ftr_refs={ftr}')
print()
for i in range(17, 25):
    if i >= len(body): break
    el  = body[i]
    tag = el.tag.split('}')[-1]
    if tag != 'p':
        print(f'[{i}] <{tag}>')
        continue
    txt   = ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()[:40]
    ppPr  = el.find('{%s}pPr' % W)
    sp    = ppPr.find('{%s}sectPr' % W) if ppPr is not None else None
    flags = ['SECT'] if sp is not None else []
    if any(br.get('{%s}type' % W) == 'page' for br in el.iter('{%s}br' % W)):
        flags.append('pgBr')
    label = repr(txt) if txt else '(empty)'
    print(f'[{i}] {label} {flags}')
