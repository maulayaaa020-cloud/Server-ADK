import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 11_p3.docx"
with zipfile.ZipFile(path) as z:
    # footer6.xml = first-page footer of section[00] (cover)
    xml = z.read("word/footer6.xml")
froot = etree.fromstring(xml)
paras = froot.findall(".//{%s}p" % W)
print("footer6.xml (first-page footer cover) paragraphs: %d" % len(paras))
for i, p in enumerate(paras):
    txt = "".join(t.text or "" for t in p.iter("{%s}t" % W))
    children = list(p)
    print("  para[%d]: %d children, text=%r" % (i, len(children), txt[:60]))
