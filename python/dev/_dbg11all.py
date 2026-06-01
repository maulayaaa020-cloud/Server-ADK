import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 11_p3.docx"
with zipfile.ZipFile(path) as z:
    rels      = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read("word/_rels/document.xml.rels"))}
    names     = set(z.namelist())
    hf_data   = {n: z.read(n) for n in names if n.startswith("word/") and (
                 n[5:].startswith("footer") or n[5:].startswith("header")) and n.endswith(".xml")}
    sects     = list(etree.fromstring(z.read("word/document.xml"))
                     .find(".//{%s}body" % W).iter("{%s}sectPr" % W))

for i, sp in enumerate(sects):
    hasTpg = sp.find("{%s}titlePg" % W) is not None
    pn     = sp.find("{%s}pgNumType" % W)
    pns    = {k.split("}")[-1]:v for k,v in pn.attrib.items()} if pn is not None else {}
    print("[%02d] titlePg=%-5s pgNum=%s" % (i, hasTpg, pns))
    for e in sp:
        tag = e.tag.split("}")[-1]
        if "Reference" not in tag: continue
        rtype = e.get("{%s}type" % W)
        rid   = e.get("{%s}id" % R)
        if not rid or rid not in rels: continue
        target = "word/" + rels[rid]
        if target not in hf_data: continue
        froot  = etree.fromstring(hf_data[target])
        hasFld = froot.find(".//{%s}fldChar" % W) is not None
        txt    = "".join(t.text or "" for t in froot.iter("{%s}t" % W))[:30]
        print("  %s/%s -> %s hasFld=%-5s %r" % (tag, rtype, rels[rid], hasFld, txt))
