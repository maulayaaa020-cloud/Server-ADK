import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

# Proses input dan jalankan strip secara manual tanpa try/except
import main as m_unused

path_in  = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 19.docx'
path_out = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 19_dbg.docx'

doc  = Document(path_in)
proc = DocProcessor(doc, 'Times New Roman', 12)

# Jalankan strip langsung — tanpa try/except agar error terlihat
body     = doc.element.body
children = list(body)
to_remove = set()

for i, elem in enumerate(children):
    if not elem.tag.endswith('}p'):
        continue
    has_page_br = any(
        br.get(qn('w:type'), '') == 'page'
        for br in elem.findall('.//' + qn('w:br'))
    )
    has_sect_br = DocProcessor._has_sectPr(elem)
    if not (has_page_br or has_sect_br):
        continue

    j = i - 1
    collected = []
    while j >= 0 and j not in to_remove:
        prev = children[j]
        if not prev.tag.endswith('}p'):
            j -= 1
            continue
        is_empty = DocProcessor._p_is_visually_empty(prev)
        has_sect = DocProcessor._has_sectPr(prev)
        if not is_empty:
            break
        if has_sect:
            break
        collected.append(j)
        to_remove.add(j)
        j -= 1

    if collected:
        txt = ''.join(t.text or '' for t in elem.iter(qn('w:t'))).strip()[:30]
        print(f"Break di children[{i}] ({repr(txt)}): akan hapus {sorted(collected)}")

print(f"\nTotal to_remove: {len(to_remove)}: {sorted(to_remove)}")

# Hapus dan cek error
errors = []
for idx in sorted(to_remove, reverse=True):
    elem_rm = children[idx]
    parent = elem_rm.getparent()
    if parent is None:
        errors.append(f"children[{idx}]: parent=None!")
        continue
    if parent is not body:
        errors.append(f"children[{idx}]: parent bukan body! ({parent.tag})")
        continue
    body.remove(elem_rm)

if errors:
    print("\nERRORS:")
    for e in errors:
        print(f"  {e}")
else:
    print(f"\nSemua dihapus tanpa error")

print(f"Body children setelah hapus: {len(list(body))}")
