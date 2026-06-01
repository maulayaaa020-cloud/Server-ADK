import zipfile, os, glob
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

hasil_dir = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil"
for path in sorted(glob.glob(os.path.join(hasil_dir, "*.docx")))[:12]:
    name = os.path.basename(path)
    with zipfile.ZipFile(path) as z:
        rels  = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read("word/_rels/document.xml.rels"))}
        names = set(z.namelist())
        doc   = z.read("word/document.xml")
        foot_data = {}
        for n in names:
            if n.startswith("word/footer") and n.endswith(".xml"):
                foot_data[n] = z.read(n)
    sects   = list(etree.fromstring(doc).find(".//{%s}body"%W).iter("{%s}sectPr"%W))
    summary = []
    for i, sp in enumerate(sects[:3]):
        tp    = sp.find("{%s}type"%W)
        stype = tp.get("{%s}val"%W) if tp is not None else "next"
        pn    = sp.find("{%s}pgNumType"%W)
        pns   = {k.split("}")[-1]:v for k,v in pn.attrib.items()} if pn is not None else {}
        fref  = next((e.get("{%s}id"%R) for e in sp
                      if "footerReference" in e.tag and e.get("{%s}type"%W)=="default"), None)
        hasFld = False
        if fref and fref in rels:
            target = "word/"+rels[fref]
            if target in foot_data:
                hasFld = etree.fromstring(foot_data[target]).find(".//{%s}fldChar"%W) is not None
        summary.append("s%d=%s pgNum=%s fld=%s" % (i, stype[:4], pns, hasFld))
    print("%-22s %s" % (name[:22], " | ".join(summary)))
