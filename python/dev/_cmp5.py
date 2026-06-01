import zipfile
from lxml import etree

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def check(label, path, sec_idx=1):
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        rels_xml = z.read("word/_rels/document.xml.rels")
        doc_xml  = z.read("word/document.xml")
        rels_root = etree.fromstring(rels_xml)
        rels = {r.get("Id"): r.get("Target") for r in rels_root}
        doc_root = etree.fromstring(doc_xml)
        body = doc_root.find(".//{%s}body" % W)
        sects = list(body.iter("{%s}sectPr" % W))
        sp = sects[sec_idx]
        # Check titlePg
        hasTpg = sp.find("{%s}titlePg" % W) is not None
        pn = sp.find("{%s}pgNumType" % W)
        pn_s = {k.split("}")[-1]:v for k,v in pn.attrib.items()} if pn is not None else {}
        refs = []
        for e in sp:
            short = e.tag.split("}")[-1]
            if "Reference" in short:
                rtype = e.get("{%s}type" % W)
                rid   = e.get("{%s}id" % R)
                refs.append((short, rtype, rid))
        print("\n%s  sec[%d]  titlePg=%s  pgNum=%s" % (label, sec_idx, hasTpg, pn_s))
        print("  refs:", refs)
        for short, rtype, rid in refs:
            if rid and rid in rels:
                target = "word/" + rels[rid]
                if target in names:
                    fxml = z.read(target)
                    froot = etree.fromstring(fxml)
                    hasFld = froot.find(".//{%s}fldChar" % W) is not None
                    # check titlePg in footer itself? no, check rPr color
                    runs = froot.findall(".//{%s}r" % W)
                    for run in runs:
                        rpr = run.find("{%s}rPr" % W)
                        if rpr is not None:
                            color = rpr.find("{%s}color" % W)
                            if color is not None:
                                cval = color.get("{%s}val" % W)
                                print("  WARN %s color=%s" % (short+"/"+rtype, cval))
                    print("  %s/%s -> %s  hasField=%s" % (short, rtype, rels[rid], hasFld))

C1 = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 5_p3.docx"
C2 = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 5_p3.docx"
check("COVER-1", C1)
check("COVER-2", C2)
