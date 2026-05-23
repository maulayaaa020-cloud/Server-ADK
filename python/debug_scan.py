import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from utils import DocProcessor, is_toc_heading, is_toc_entry, is_bab_heading, is_false_bab, BAB_HEAD_RE

doc  = Document(sys.argv[1])
proc = DocProcessor(doc, "Times New Roman", 12)

# Patch scan_zones dengan logging
all_paras = list(doc.paragraphs)
print(f"Total paragraf: {len(all_paras)}")

roman_start_p      = None
bab_p_list         = []
lampiran_found     = False
last_bab_para_idx  = None
found_numbered_bab = False
seen_bab_info      = {}
inside_toc         = False
toc_start_idx      = -1

import re
from docx.oxml.ns import qn

def has_sectPr(p_elem):
    pPr = p_elem.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:sectPr')) is not None

for para_idx, para in enumerate(all_paras):
    text  = para.text.strip()
    lower = text.lower()

    if is_toc_heading(lower):
        inside_toc    = True
        toc_start_idx = para_idx
        if roman_start_p is None:
            roman_start_p = para._p
        print(f"[{para_idx:3d}] TOC_HEADING: {repr(text[:50])}")
        continue

    if inside_toc:
        prev_has_sect = any(
            has_sectPr(all_paras[j]._p)
            for j in range(max(toc_start_idx + 1, para_idx - 5), para_idx)
        )
        if prev_has_sect or proc._para_has_page_break_before(para):
            reason = "SECT_BREAK" if prev_has_sect else "PAGE_BREAK"
            print(f"[{para_idx:3d}] EXIT_TOC via {reason}: {repr(text[:50])}")
            inside_toc = False
        elif not text or is_toc_entry(text):
            continue
        else:
            still_in_toc = False
            for _lk in range(para_idx + 1, min(para_idx + 9, len(all_paras))):
                _lp = all_paras[_lk]
                _ls = (_lp.style.name.lower() if _lp.style else "")
                _lt = _lp.text.strip()
                if 'toc' in _ls:
                    still_in_toc = True; break
                if _lt and is_toc_entry(_lt):
                    still_in_toc = True; break
                if _lt and is_bab_heading(_lt):
                    _lp_heading = bool(re.search(r'heading', _ls))
                    _lp_has_brk = proc._para_has_page_break_before(_lp)
                    if _lp_heading or _lp_has_brk:
                        print(f"[{para_idx:3d}] LOOKAHEAD EXIT (real BAB at {_lk}): {repr(text[:40])}")
                        break
                    print(f"[{para_idx:3d}] LOOKAHEAD STAY (fake BAB at {_lk}): {repr(text[:40])}")
                    still_in_toc = True; break
            if still_in_toc:
                continue
            print(f"[{para_idx:3d}] EXIT_TOC via lookahead no-signal: {repr(text[:50])}")
            inside_toc = False

    if not text:
        continue

    if is_bab_heading(text) and not is_false_bab(para):
        is_numbered_bab = BAB_HEAD_RE.match(text) is not None
        if not is_numbered_bab and not found_numbered_bab:
            continue
        _style_lower  = (para.style.name.lower() if para.style else "")
        _is_heading   = bool(re.search(r'heading', _style_lower))
        _has_brk_para = proc._para_has_page_break_before(para)
        if is_numbered_bab:
            m_num = BAB_HEAD_RE.match(text)
            bab_num_key = m_num.group(2).strip().lower() if m_num else None
            if bab_num_key and bab_num_key in seen_bab_info:
                old = seen_bab_info[bab_num_key]
                replaced = (_has_brk_para or _is_heading) and not (old['has_break'] or old['is_heading'])
                print(f"[{para_idx:3d}] DUPLICATE '{bab_num_key}': replaced={replaced} (new brk={_has_brk_para},hdg={_is_heading} | old brk={old['has_break']},hdg={old['is_heading']})")
                if replaced:
                    bab_p_list[old['idx']] = para._p
                    old['has_break'] = _has_brk_para
                    old['is_heading'] = _is_heading
                continue
        else:
            bab_num_key = None

        print(f"[{para_idx:3d}] BAB_ADD: {repr(text[:50])} (brk={_has_brk_para},hdg={_is_heading})")
        bab_p_list.append(para._p)
        if bab_num_key:
            seen_bab_info[bab_num_key] = {'idx': len(bab_p_list)-1, 'has_break': _has_brk_para, 'is_heading': _is_heading}
        last_bab_para_idx = para_idx
        found_numbered_bab = found_numbered_bab or is_numbered_bab

print("\n=== RESULT ===")
for i, p in enumerate(bab_p_list):
    from docx.oxml.ns import qn as Q
    t = ''.join(x.text or '' for x in p.iter(Q('w:t'))).strip()
    print(f"  [{i}] {t[:60]}")
