import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 5_p3.docx"
with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(".//{%s}body" % W)
sects = list(body.iter("{%s}sectPr" % W))

for i, sp in enumerate(sects[:4]):
    print("=== sectPr[%d] ===" % i)
    for child in sp:
        tag  = child.tag.split("}")[-1]
        attr = {k.split("}")[-1]: v for k, v in child.attrib.items()}
        print("  <%s %s>" % (tag, attr))
    print()
