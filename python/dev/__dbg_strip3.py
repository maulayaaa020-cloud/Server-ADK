import sys, os
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')

# Jalankan main.py dengan output ke file temp, lalu baca dan test strip
import subprocess, json
path_in  = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
path_tmp = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 19_nostip.docx'

# Baca file hasil yang sudah ada (hasil terbaru)
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

# Buka hasil file yang sudah ada
path_hasil = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 19.docx'
doc  = Document(path_hasil)
proc = DocProcessor(doc, 'Times New Roman', 12)

body     = doc.element.body
children = list(body)
to_remove = set()

print(f"Total body children di hasil: {len(children)}")
print()

for i, elem in enumerate(children):
    if not elem.tag.endswith('}p'):
        continue
    has_page_br = any(br.get(qn('w:type'), '') == 'page'
                      for br in elem.findall('.//' + qn('w:br')))
    has_sect_br = DocProcessor._has_sectPr(elem)
    if not (has_page_br or has_sect_br):
        continue

    collected = []
    j = i - 1
    while j >= 0 and j not in to_remove:
        prev = children[j]
        if not prev.tag.endswith('}p'):
            j -= 1; continue
        is_empty = DocProcessor._p_is_visually_empty(prev)
        has_sect = DocProcessor._has_sectPr(prev)
        if not is_empty:
            txt = ''.join(t.text or '' for t in prev.iter(qn('w:t'))).strip()[:30]
            print(f"  STOP at j={j}: not visually empty, text={repr(txt)}")
            break
        if has_sect:
            print(f"  STOP at j={j}: has sectPr")
            break
        collected.append(j)
        to_remove.add(j)
        j -= 1

    if collected:
        hdr = ''.join(t.text or '' for t in elem.iter(qn('w:t'))).strip()[:30]
        print(f"Break[{i}] {repr(hdr)}: hapus {sorted(collected)}")

print()
print(f"to_remove total: {sorted(to_remove)}")

# Cek parent dari to_remove elements
print()
errors = []
for idx in sorted(to_remove, reverse=True):
    e = children[idx]
    p = e.getparent()
    if p is None:
        errors.append(f"[{idx}] parent=None")
    elif p is not body:
        errors.append(f"[{idx}] parent bukan body: {p.tag}")

if errors:
    print("ERRORS:")
    for err in errors:
        print(f"  {err}")
else:
    print("Semua element to_remove adalah direct children of body")
