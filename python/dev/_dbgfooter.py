import zipfile
from lxml import etree

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 5_p3.docx"
with zipfile.ZipFile(path) as z:
    names = set(z.namelist())
    rels_xml = z.read("word/_rels/document.xml.rels")
    doc_xml  = z.read("word/document.xml")
    
    rels_root = etree.fromstring(rels_xml)
    rels = {r.get("Id"): r.get("Target") for r in rels_root}
    
    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(".//{%s}body" % W)
    sects = list(body.iter("{%s}sectPr" % W))
    
    for i, sp in enumerate(sects[:6]):
        refs = []
        for e in sp:
            short = e.tag.split("}")[-1]
            if "Reference" in short:
                rtype = e.get("{%s}type" % W)
                rid   = e.get("{%s}id" % R)
                refs.append((short, rtype, rid))
        print("[%02d] refs=%s" % (i, [(t,rt,rid) for t,rt,rid in refs]))
        for short, rtype, rid in refs:
            if rid and rid in rels:
                target = "word/" + rels[rid]
                if target in names:
                    fxml  = z.read(target)
                    froot = etree.fromstring(fxml)
                    txt   = "".join(t.text or "" for t in froot.iter("{%s}t" % W))
                    hasFld = froot.find(".//{%s}fldChar" % W) is not None
                    print("     %s/%s rid=%s -> %s  text=%r hasField=%s" % (short, rtype, rid, rels[rid], txt[:40], hasFld))
            elif rid:
                print("     %s/%s rid=%s NOT IN RELS" % (short, rtype, rid))
