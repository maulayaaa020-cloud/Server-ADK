import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
OUT = r'C:\Users\farizal\Downloads\GEMA_c4_semb_lamprn.docx'

doc  = Document(OUT)
body = list(doc.element.body)

print(f'Total body elements: {len(body)}')
print(f'Total sections: {len(doc.sections)}')

print('\n=== body[1020:1055] ===')
for i in range(1020, min(1055, len(body))):
    el  = body[i]
    tag = el.tag.split('}')[-1]
    if tag != 'p':
        print(f'  [{i}] <{tag}>')
        continue
    txt  = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:60]
    pPr  = el.find(f'{{{W}}}pPr')
    sp   = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
    pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
    flags = []
    if sp is not None:
        t2 = sp.find(f'{{{W}}}type')
        flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
    if pgbr:
        flags.append('pgBr')
    print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')

print('\n=== Section headers/footers ===')
for i, sec in enumerate(doc.sections):
    pn  = sec._sectPr.find(qn('w:pgNumType'))
    fmt = pn.get(qn('w:fmt'))   if pn is not None else None
    st  = pn.get(qn('w:start')) if pn is not None else None
    # Check if header/footer have any runs (page number fields etc.)
    def has_content(part):
        try:
            for p in part.paragraphs:
                for run in p.runs:
                    if run.text.strip():
                        return True
                if p._p.findall('.//' + qn('w:fldChar')):
                    return True
            return False
        except:
            return None
    hdr_has = has_content(sec.header)
    ftr_has = has_content(sec.footer)
    print(f'  sec[{i}] fmt={fmt} start={st} hdr_has_content={hdr_has} ftr_has_content={ftr_has}')
