import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
SRC = r'C:\Users\farizal\Downloads\SKRIPSI GEMA RAMADHAN fiks.docx'

doc    = Document(SRC)
paras  = list(doc.paragraphs)
body   = list(doc.element.body)

# Find "Lampiran 1" in body
for i, el in enumerate(body):
    if el.tag.endswith('}p'):
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()
        if txt == 'Lampiran 1':
            print(f'Found "Lampiran 1" at body[{i}]')
            print('\n=== body[{0}:{1}] (30 before) ==='.format(i-30, i+5))
            for j in range(max(0, i-30), i+5):
                el2 = body[j]
                tag = el2.tag.split('}')[-1]
                if tag != 'p':
                    print(f'  [{j}] <{tag}>')
                    continue
                txt2 = ''.join(t.text or '' for t in el2.iter(f'{{{W}}}t')).strip()
                pPr  = el2.find(f'{{{W}}}pPr')
                sp   = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
                pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el2.iter(f'{{{W}}}br'))
                flags = []
                if sp is not None:
                    flags.append('SECT')
                if pgbr:
                    flags.append('pgBr')
                sty = el2.find(f'{{{W}}}pPr')
                sty = sty.find(f'{{{W}}}pStyle') if sty is not None else None
                sty_val = sty.get(qn('w:val')) if sty is not None else None
                print(f'  [{j}] style={sty_val!r} {repr(txt2) if txt2 else "(empty)"} {flags}')
            break

# Also find in all_paras
print('\n=== all_paras scan for single-letter / LAMPIRAN ===')
for i, para in enumerate(paras):
    t = para.text.strip()
    if len(t) == 1 and t.isalpha():
        print(f'  para[{i}] body_el={body.index(para._p) if para._p in body else "nested"} style={para.style.name!r} char={repr(t)}')
    elif t.lower() in ('lampiran', 'lampiran 1', 'lampiran 2'):
        pPr = para._p.find(f'{{{W}}}pPr')
        sp  = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
        pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in para._p.iter(f'{{{W}}}br'))
        print(f'  para[{i}] style={para.style.name!r} text={repr(t)} sect={sp is not None} pgBr={pgbr}')
