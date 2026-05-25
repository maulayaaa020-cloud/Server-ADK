"""Check cover 1 dimulai dari ii - correct check: lowerRoman:1 cover, then lowerRoman:None roman"""
import os, sys, json, subprocess
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn

ROOT    = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
FOLDER  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii')
HASIL   = os.path.join(FOLDER, 'hasil')
MAIN_PY = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'main.py')
PYTHON  = r'D:\Freelaces\Server\python.exe'
ARGS    = ['paket3','Times New Roman','12 pt','Ya','Tengah Bawah','Tengah Bawah','Kanan Atas','ii','Tidak','Tidak','1']
W       = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def read_pns(path):
    doc = Document(path)
    r = []
    for sec in doc.sections:
        pn = sec._sectPr.find(qn('w:pgNumType'))
        r.append({'fmt': pn.get(qn('w:fmt'),'decimal') if pn is not None else 'decimal',
                  'start': pn.get(qn('w:start'),None) if pn is not None else None})
    return r

def check(pns):
    # Cover 1, dimulai_dari='ii': cover=lowerRoman:1, roman=lowerRoman:None, decimal after
    lr1 = next((i for i,p in enumerate(pns) if p['fmt']=='lowerRoman' and p['start']=='1'), None)
    if lr1 is None: return False, 'no lowerRoman:1 cover'
    # Roman section: lowerRoman with None start, right after cover
    lr_none = next((i for i,p in enumerate(pns) if i > lr1 and p['fmt']=='lowerRoman' and p['start'] is None), None)
    if lr_none is None: return False, 'no lowerRoman:None roman section after cover'
    dec = next((i for i,p in enumerate(pns) if i > lr_none and p['fmt']=='decimal'), None)
    if dec is None: return False, 'no decimal section after roman'
    return True, "cov@sec[%d], rom@sec[%d], dec@sec[%d]" % (lr1, lr_none, dec)

os.makedirs(HASIL, exist_ok=True)
ok=bad=err=0
for fname in sorted(f for f in os.listdir(FOLDER) if f.endswith('.docx')):
    inp = os.path.join(FOLDER, fname)
    out = os.path.join(HASIL, os.path.splitext(fname)[0]+'_c1.docx')
    res = subprocess.run([PYTHON, MAIN_PY, inp, out]+ARGS, capture_output=True, text=True, encoding='utf-8')
    try: data = json.loads(res.stdout)
    except: print("  ERR  %s  %s" % (fname, (res.stderr or res.stdout or '').strip()[:80])); err+=1; continue
    if data.get('status')!='success': print("  ERR  %s  %s" % (fname, data.get('message','')[:80])); err+=1; continue
    pns = read_pns(out)
    ok_f, reason = check(pns)
    tag = 'OK ' if ok_f else 'BAD'
    if ok_f:
        print("  %s  %s  %s" % (tag, fname, reason))
    else:
        print("  %s  %s  %s\n        %s" % (tag, fname, reason, str([p['fmt']+':'+str(p['start']) for p in pns])))
    if ok_f: ok+=1
    else: bad+=1
print("\n%s\n  %d OK | %d BAD | %d ERR\n%s\nOutput: %s" % ('='*50, ok, bad, err, '='*50, HASIL))
