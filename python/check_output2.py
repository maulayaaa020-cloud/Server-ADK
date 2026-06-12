import zipfile
from lxml import etree

path = r"C:\Users\farizal\Downloads\Test Dafis - Output.docx"
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")

tree = etree.fromstring(xml)
body = tree.find(f"{{{ns}}}body")

# Get DIRECT children of body only
children = list(body)
print(f"Total direct body children: {len(children)}")

for i, child in enumerate(children):
    tag = child.tag.split("}")[-1]
    if tag == "p":
        runs = child.findall(f".//{{{ns}}}t")
        text = "".join(r.text or "" for r in runs).strip()[:60]
        pPr = child.find(f"{{{ns}}}pPr")
        style = ""
        if pPr is not None:
            pStyle = pPr.find(f"{{{ns}}}pStyle")
            if pStyle is not None:
                style = pStyle.get(f"{{{ns}}}val", "")
        sect = child.find(f".//{{{ns}}}sectPr")
        fld = child.find(f".//{{{ns}}}fldChar")
        instr = child.find(f".//{{{ns}}}instrText")
        flags = ""
        if sect is not None: flags += " [SECT]"
        if fld is not None: flags += " [FLD]"
        if instr is not None: flags += f" [INSTR:{(instr.text or '')[:30]}]"
        print(f"  [{i}] p style={style:15} {flags} '{text}'")
    else:
        print(f"  [{i}] {tag}")

    if i > 50:
        print("...")
        break
