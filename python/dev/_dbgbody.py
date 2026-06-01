import zipfile
from lxml import etree

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WPD = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 5_p3.docx"
with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(".//{%s}body" % W)
children = list(body)

# Find section breaks
sect_idxs = []
for i, el in enumerate(children):
    if el.tag == "{%s}p" % W:
        if el.find(".//{%s}sectPr" % W) is not None:
            sect_idxs.append(i)

# Show paragraphs around sectPr[0] and sectPr[1]
print("sectPr indices: %s" % sect_idxs[:4])
for si in sect_idxs[:2]:
    print("\n--- around sectPr[%d] (paragraph idx=%d) ---" % (sect_idxs.index(si), si))
    for i in range(max(0, si-3), min(len(children), si+4)):
        el = children[i]
        tag = el.tag.split("}")[-1]
        txt = "".join(t.text or "" for t in el.iter("{%s}t" % W)).strip()
        hasSect = el.find(".//{%s}sectPr" % W) is not None
        hasPgBr = any(br.get("{%s}type" % W) == "page" for br in el.iter("{%s}br" % W))
        marks = ("[SECT]" if hasSect else "") + ("[PGBR]" if hasPgBr else "")
        print("  [%03d] %-16s %s" % (i, marks, txt[:70]))
