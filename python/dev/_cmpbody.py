import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def check_body(label, path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(".//{%s}body" % W)
    children = list(body)
    sect_idxs = [i for i, el in enumerate(children) 
                 if el.tag == "{%s}p" % W and el.find(".//{%s}sectPr" % W) is not None]
    print("\n%s  section breaks at: %s" % (label, sect_idxs[:4]))
    si = sect_idxs[0]
    print("  Around sectPr[0] (para %d):" % si)
    for i in range(max(0, si-3), min(len(children), si+5)):
        el = children[i]
        txt = "".join(t.text or "" for t in el.iter("{%s}t" % W)).strip()[:60]
        hasSect = el.find(".//{%s}sectPr" % W) is not None
        hasPgBr = any(br.get("{%s}type" % W) == "page" for br in el.iter("{%s}br" % W))
        print("  [%03d] %s%s  %r" % (i, "[SECT]" if hasSect else "      ", "[PGBR]" if hasPgBr else "      ", txt))

C1 = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 5_p3.docx"
C2 = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 5_p3.docx"
check_body("COVER-1", C1)
check_body("COVER-2", C2)
