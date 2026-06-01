import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RI = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

SRC = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 16.docx'

def show_body(doc, label, start=0, end=60):
    body = list(doc.element.body)
    print(f'\n[{label}] body[{start}:{end}]:')
    for i in range(start, min(end, len(body))):
        el = body[i]
        tag = el.tag.split('}')[-1]
        if tag != 'p':
            print(f'  [{i}] <{tag}>')
            continue
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:40]
        pPr = el.find(f'{{{W}}}pPr')
        sp = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
        flags = []
        if sp is not None:
            t2 = sp.find(f'{{{W}}}type')
            flags.append(f'SECT({t2.get(qn("w:val")) if t2 is not None else "nextPage"})')
        # check pgBr
        has_pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
        if has_pgbr:
            flags.append('pgBr')
        print(f'  [{i}] {repr(txt) if txt else "(empty)"} {flags}')

def show_sections(doc, label):
    print(f'\n[{label}] sections:')
    for i, sec in enumerate(doc.sections):
        sp = sec._sectPr
        t = sp.find(qn('w:type'))
        tval = t.get(qn('w:val')) if t is not None else 'nextPage'
        frefs = [(fr.get(qn('w:type')), fr.get(RI)) for fr in sp.findall(qn('w:footerReference'))]
        hrefs = [(hr.get(qn('w:type')), hr.get(RI)) for hr in sp.findall(qn('w:headerReference'))]
        body = list(doc.element.body)
        para_txt = ''
        for el in body:
            if not el.tag.endswith('}p'): continue
            ppPr = el.find(f'{{{W}}}pPr')
            if ppPr is not None and ppPr.find(f'{{{W}}}sectPr') is sp:
                para_txt = ''.join(t2.text or '' for t2 in el.iter(f'{{{W}}}t')).strip()[:25]
                break
        loc = f' @ "{para_txt}"' if para_txt else ' @ [body]'
        print(f'  sec[{i}] {tval:12} hdr={hrefs} ftr={frefs}{loc}')

doc  = Document(SRC)
proc = DocProcessor(doc, 'Times New Roman', 12)

print('=== STEP 0: initial body structure ===')
show_body(doc, 'initial', 0, 50)
show_sections(doc, 'initial')

print('\n=== STEP 1: purge_all_headers_footers ===')
proc.purge_all_headers_footers()
show_body(doc, 'after purge', 0, 50)

print('\n=== STEP 2: scan_zones ===')
rsp, bab = proc.scan_zones()
body = list(doc.element.body)
print(f'  roman_start_p: [{body.index(rsp)}] "{DocProcessor._p_text(rsp)[:40]}"')
print(f'  bab_p_list: {len(bab)} items')
for b in bab:
    print(f'    [{body.index(b)}] "{DocProcessor._p_text(b)[:30]}"')

print('\n=== STEP 3: advance_roman_start ===')
new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
body = list(doc.element.body)
new_idx = body.index(new_rsp) if new_rsp in body else -1
print(f'  new_rsp=[{new_idx}] "{DocProcessor._p_text(new_rsp)[:40]}" exact={use_exact}')
if new_rsp is not rsp: rsp = new_rsp

print('\n=== STEP 4: body around new rsp ===')
show_body(doc, 'pre-insert', max(0, new_idx-3), new_idx+5)

print('\n=== STEP 5: insert_breaks ===')
rsp = proc.insert_breaks(rsp, bab, exact_roman_start=use_exact)
body = list(doc.element.body)
rsp_idx = body.index(rsp) if rsp in body else -1
print(f'  rsp after=[{rsp_idx}] "{DocProcessor._p_text(rsp)[:40]}"')
show_sections(doc, 'after insert_breaks')
show_body(doc, 'after insert_breaks', 0, 55)
