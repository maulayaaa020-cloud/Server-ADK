"""Check cover 2 dimulai dari iii"""
import os, sys, json, subprocess
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')))
from docx import Document
from docx.oxml.ns import qn

ROOT    = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
FOLDER  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii')
HASIL   = os.path.join(FOLDER, 'hasil')
MAIN_PY = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'main.py')
PYTHON  = r'D:\Freelaces\Server\python.exe'
ARGS    = ['paket3','Times New Roman','12 pt','Ya','Tengah Bawah','Tengah Bawah','Kanan Atas','iii','Tidak','Tidak','2']
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
    lr3 = next((i for i,p in enumerate(pns) if p['fmt']=='lowerRoman' and p['start']=='3'), None)
    if lr3 is None: return False, 'no lowerRoman:3'
    dec = next((i for i in range(lr3+1,len(pns)) if pns[i]['fmt']=='decimal'), None)
    if dec is None: return False, 'no decimal after lowerRoman:3'
    return True, f"lr3@sec[{lr3}], dec@sec[{dec}]"

os.makedirs(HASIL, exist_ok=True)
ok=bad=err=0
for fname in sorted(f for f in os.listdir(FOLDER) if f.endswith('.docx')):
    inp = os.path.join(FOLDER, fname)
    out = os.path.join(HASIL, os.path.splitext(fname)[0]+'_v2.docx')
    res = subprocess.run([PYTHON, MAIN_PY, inp, out]+ARGS, capture_output=True, text=True, encoding='utf-8')
    try: data = json.loads(res.stdout)
    except: print(f"  ERR  {fname}  {(res.stderr or res.stdout or '').strip()[:80]}"); err+=1; continue
    if data.get('status')!='success': print(f"  ERR  {fname}  {data.get('message','')[:80]}"); err+=1; continue
    pns = read_pns(out)
    ok_f, reason = check(pns)
    tag = 'OK ' if ok_f else 'BAD'
    (print(f"  {tag}  {fname}  {reason}") if ok_f else
     print(f"  {tag}  {fname}  {reason}\n        {[p['fmt']+':'+str(p['start']) for p in pns]}"))
    if ok_f: ok+=1
    else: bad+=1
print(f"\n{'='*50}\n  {ok} OK | {bad} BAD | {err} ERR\n{'='*50}\nOutput: {HASIL}")
