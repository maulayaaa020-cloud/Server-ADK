import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RI = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

def check(path, label):
    doc = Document(path)
    body = list(doc.element.body)
    print(f'\n=== {label} ===')
    print(f'Total sections: {len(list(doc.sections))}')
    for i, sec in enumerate(doc.sections):
        sp = sec._sectPr
        t = sp.find(qn('w:type'))
        tval = t.get(qn('w:val')) if t is not None else 'nextPage'
        pn = sp.find(qn('w:pgNumType'))
        fmt   = pn.get(qn('w:fmt'))   if pn is not None else None
        start = pn.get(qn('w:start')) if pn is not None else None
        frefs = [(fr.get(qn('w:type')), fr.get(RI)) for fr in sp.findall(qn('w:footerReference'))]
        hrefs = [(hr.get(qn('w:type')), hr.get(RI)) for hr in sp.findall(qn('w:headerReference'))]
        # first_content
        first = ''
        in_sec = False
        for j, el in enumerate(body):
            if not el.tag.endswith('}p'): continue
            ppPr = el.find(f'{{{W}}}pPr')
            prev_sp = ppPr.find(f'{{{W}}}sectPr') if ppPr is not None else None
            if prev_sp is not None:
                # end of previous section
                if i == 0 and not in_sec:
                    in_sec = True
                    continue
                if in_sec:
                    break
            if i == 0 and not in_sec:
                txt = ''.join(t2.text or '' for t2 in el.iter(f'{{{W}}}t')).strip()
                if txt and not first:
                    first = txt[:30]
        # simpler: use sections boundary
        first = ''
        para_txt = ''
        for el in body:
            if not el.tag.endswith('}p'): continue
            ppPr = el.find(f'{{{W}}}pPr')
            if ppPr is not None and ppPr.find(f'{{{W}}}sectPr') is sp:
                para_txt = ''.join(t2.text or '' for t2 in el.iter(f'{{{W}}}t')).strip()[:30]
                break
        loc = f'@ para:"{para_txt}"' if para_txt else '@ [body]'
        print(f'  sec[{i}] {tval:12} fmt={fmt} start={start} hdr={len(hrefs)} ftr={len(frefs)} {loc}')

    print(f'\nBody [0:55]:')
    for i in range(min(55, len(body))):
        el = body[i]
        tag = el.tag.split('}')[-1]
        if tag == 'tbl':
            print(f'  [{i}] <tbl>')
            continue
        if tag != 'p':
            print(f'  [{i}] <{tag}>')
            continue
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:40]
        ppPr = el.find(f'{{{W}}}pPr')
        sp2 = ppPr.find(f'{{{W}}}sectPr') if ppPr is not None else None
        flags = []
        if sp2 is not None:
            t2 = sp2.find(f'{{{W}}}type')
            flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
        if any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br')):
            flags.append('pgBr')
        print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')

check(
    r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 10_c2.docx',
    'Docx 10 c2'
)
