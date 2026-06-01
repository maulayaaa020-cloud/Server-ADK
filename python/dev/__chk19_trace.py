import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from utils import DocProcessor, is_bab_heading, is_false_bab, is_roman_start, is_toc_heading, is_toc_entry, _has_toc_field
import re

path = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
doc  = Document(path)
BAB_HEAD_RE = re.compile(
    r'^\s*(bab|chapter)\s+([IVXLCDM]+|\d+)\b', re.IGNORECASE
)

proc     = DocProcessor(doc, 'Times New Roman', 12)
all_paras = list(doc.paragraphs)

inside_toc        = False
toc_start_idx     = -1
found_numbered_bab = False
roman_start_p     = None

print("Tracing scan_zones — fokus area sekitar BAB I (para 60-135) dan BAB IV (para 245-260)\n")

TARGET_RANGES = list(range(60, 135)) + list(range(245, 262))

for para_idx, para in enumerate(all_paras):
    text  = para.text.strip()
    lower = text.lower()
    in_range = para_idx in TARGET_RANGES

    # TOC heading check
    if is_toc_heading(lower) or _has_toc_field(para._p):
        if not inside_toc:
            toc_start_idx = para_idx
        inside_toc = True
        if roman_start_p is None:
            roman_start_p = para._p
        if in_range:
            print(f"[{para_idx:3d}] TOC_HEADING  inside_toc=True  text={repr(text[:50])}")
        continue

    if inside_toc:
        prev_has_break = proc._has_page_break_before(
            all_paras, max(toc_start_idx + 1, para_idx - 15), para_idx)
        has_brk_para = proc._para_has_page_break_before(para)
        if prev_has_break or has_brk_para:
            inside_toc = False
            if in_range:
                print(f"[{para_idx:3d}] EXIT_TOC(break)  text={repr(text[:50])}")
        elif not text or is_toc_entry(text) or (
            re.match(r'^\d+\.\d', text) and
            not re.search(r'heading', para.style.name.lower() if para.style else '')
        ):
            if in_range and text:
                print(f"[{para_idx:3d}] TOC_ENTRY(skip)  text={repr(text[:50])}")
            continue
        else:
            _cur_style = (para.style.name.lower() if para.style else "")
            if (is_bab_heading(text)
                    and not re.search(r'heading', _cur_style)
                    and not proc._para_has_page_break_before(para)):
                if in_range:
                    print(f"[{para_idx:3d}] BAB_AS_TOC(skip) inside_toc=True  text={repr(text[:50])}")
                continue
            still_in_toc = False
            for _lk in range(para_idx + 1, min(para_idx + 9, len(all_paras))):
                _lp = all_paras[_lk]
                _ls = (_lp.style.name.lower() if _lp.style else "")
                _lt = _lp.text.strip()
                if 'toc' in _ls or (text and is_toc_entry(_lt)):
                    still_in_toc = True; break
                if _lt and re.match(r'^\d+\d', _lt) and not re.search(r'heading', _ls):
                    still_in_toc = True; break
            if still_in_toc:
                if in_range:
                    print(f"[{para_idx:3d}] STILL_TOC(skip)  text={repr(text[:50])}")
                continue
            inside_toc = False
            if in_range:
                print(f"[{para_idx:3d}] EXIT_TOC(content)  text={repr(text[:50])}")

    if not text:
        continue

    if roman_start_p is None and is_roman_start(text):
        roman_start_p = para._p
        if in_range:
            print(f"[{para_idx:3d}] ROMAN_START  text={repr(text[:50])}")

    if is_bab_heading(text) and not is_false_bab(para):
        is_numbered = BAB_HEAD_RE.match(text) is not None
        if not is_numbered and not found_numbered_bab:
            if in_range:
                print(f"[{para_idx:3d}] BAB_SKIP(no_numbered_yet)  text={repr(text[:50])}")
            continue
        if in_range:
            print(f"[{para_idx:3d}] BAB_CANDIDATE  is_numbered={is_numbered}  found_numbered={found_numbered_bab}  text={repr(text[:50])}")
        if is_numbered:
            found_numbered_bab = True
    elif in_range and text and is_bab_heading(text):
        print(f"[{para_idx:3d}] FALSE_BAB(skip)  text={repr(text[:50])}")
