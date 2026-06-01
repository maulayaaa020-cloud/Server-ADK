import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 11_p3.docx"
with zipfile.ZipFile(path) as z:
    names = set(z.namelist())
    rels  = {r.get("Id"): r.get("Target")
             for r in etree.fromstring(z.read("word/_rels/document.xml.rels"))}
    sects = list(etree.fromstring(z.read("word/document.xml"))
                      .find(".//{%s}body" % W)
                      .iter("{%s}sectPr" % W))
    for i, sp in enumerate(sects[:3]):
        hasTpg = sp.find("{%s}titlePg" % W) is not None
        pn     = sp.find("{%s}pgNumType" % W)
        pns    = {k.split("}")[-1]:v for k,v in pn.attrib.items()} if pn is not None else {}
        refs   = [(e.get("{%s}type"%W), e.get("{%s}id"%R))
                  for e in sp if "Reference" in e.tag.split("}")[-1]]
        print("[%02d] titlePg=%-5s pgNum=%s" % (i, hasTpg, pns))
        for rt, rid in refs:
            if rid and rid in rels:
                target = "word/" + rels[rid]
                if target in names:
                    froot  = etree.fromstring(z.read(target))
                    hasFld = froot.find(".//{%s}fldChar" % W) is not None
                    print("  %s/%s -> %s  hasFld=%s" % (rt, rid, rels[rid], hasFld))
