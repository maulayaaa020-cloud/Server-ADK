import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
src = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii'

for n in [1, 5, 7, 9, 10]:
    path = os.path.join(src, f'Docx {n}.docx')
    doc  = Document(path)
    proc = DocProcessor(doc, 'Times New Roman', 12)
    proc.purge_all_headers_footers()
    roman_start_p, bab_p_list = proc.scan_zones()
    body = list(doc.element.body)
    rsp_orig = body.index(roman_start_p) if roman_start_p is not None else -1

    new_rsp, use_exact = DocProcessor.advance_roman_start(doc, roman_start_p, 2)
    body2 = list(doc.element.body)
    rsp_adv = body2.index(new_rsp) if new_rsp is not None else -1

    roman_start_p = new_rsp
    roman_start_p = proc.insert_breaks(roman_start_p, bab_p_list, exact_roman_start=use_exact)

    body3 = list(doc.element.body)
    rsp_final = body3.index(roman_start_p) if roman_start_p is not None else -1

    # Simulate ensure_cover_pages counting
    body_els = body3
    rsp_idx  = rsp_final
    term_idx = -1
    for j in range(rsp_idx - 1, -1, -1):
        el = body_els[j]
        if el.tag.endswith('}p'):
            pPr = el.find('{%s}pPr' % W)
            if pPr is not None and pPr.find('{%s}sectPr' % W) is not None:
                term_idx = j
                break

    cover_breaks = 0
    for j in range(rsp_idx):
        if j == term_idx:
            continue
        el = body_els[j]
        if any(br.get('{%s}type' % W) == 'page' for br in el.iter('{%s}br' % W)):
            cover_breaks += 1
        pPr = el.find('{%s}pPr' % W)
        if pPr is not None and pPr.find('{%s}sectPr' % W) is not None:
            cover_breaks += 1

    # implicit detection (simplified)
    def _last_txt(end_idx):
        for j in range(end_idx, -1, -1):
            el = body_els[j]
            if not el.tag.endswith('}p'):
                continue
            t = ''.join(tx.text or '' for tx in el.iter('{%s}t' % W)).strip()
            if t:
                return t, j
        return None, -1

    lt, last_j = _last_txt(term_idx)
    implicit = 0
    if lt and len(lt) >= 4 and last_j >= 0:
        prev_j = last_j
        for j in range(last_j - 1, -1, -1):
            el = body_els[j]
            if not el.tag.endswith('}p'):
                continue
            t = ''.join(tx.text or '' for tx in el.iter('{%s}t' % W)).strip()
            if t == lt and prev_j - j >= 5:
                implicit += 1
                prev_j = j

    missing = (2 - 1) - (cover_breaks + implicit)
    print(f'Docx {n}: scan_rsp={rsp_orig} adv_rsp={rsp_adv} use_exact={use_exact} '
          f'final_rsp={rsp_final} term={term_idx} '
          f'cover_breaks={cover_breaks} implicit={implicit} missing={missing} '
          f'last_txt={repr(lt[:30]) if lt else None}')
