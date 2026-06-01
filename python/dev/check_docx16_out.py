import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RI = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

OUT = r'D:\Freelaces\Server\htdocs\adk\Docx 16.docx\Docx 16_p3.docx'

doc = Document(OUT)
body = list(doc.element.body)

print('=== Docx 16 output sections ===')
for i, sec in enumerate(doc.sections):
    sp = sec._sectPr
    t = sp.find(qn('w:type'))
    tval = t.get(qn('w:val')) if t is not None else 'nextPage'
    frefs = [(fr.get(qn('w:type')), fr.get(RI)) for fr in sp.findall(qn('w:footerReference'))]
    hrefs = [(hr.get(qn('w:type')), hr.get(RI)) for hr in sp.findall(qn('w:headerReference'))]
    pgNum = sp.find(qn('w:pgNumType'))
    fmt_val = pgNum.get(qn('w:fmt')) if pgNum is not None else None
    start_val = pgNum.get(qn('w:start')) if pgNum is not None else None
    # find first_content
    para_txt = ''
    for el in body:
        if not el.tag.endswith('}p'): continue
        ppPr = el.find(f'{{{W}}}pPr')
        if ppPr is not None and ppPr.find(f'{{{W}}}sectPr') is sp:
            para_txt = ''.join(t2.text or '' for t2 in el.iter(f'{{{W}}}t')).strip()[:30]
            break
    loc = f' @ para:"{para_txt}"' if para_txt else ' @ [body]'
    print(f'  sec[{i}] {tval:12} fmt={fmt_val} start={start_val} hdr={len(hrefs)} ftr={len(frefs)}{loc}')

print('\n=== Body [0:50] ===')
for i in range(min(50, len(body))):
    el = body[i]
    tag = el.tag.split('}')[-1]
    if tag == 'tbl':
        print(f'  [{i}] <tbl>')
        continue
    if tag != 'p':
        print(f'  [{i}] <{tag}>')
        continue
    txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:40]
    pPr = el.find(f'{{{W}}}pPr')
    sp2 = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
    flags = []
    if sp2 is not None:
        t2 = sp2.find(f'{{{W}}}type')
        flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
    has_pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
    if has_pgbr:
        flags.append('pgBr')
    print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')
