import sys, os, subprocess

_HERE   = os.path.dirname(os.path.abspath(__file__))
_ROOT   = os.path.normpath(os.path.join(_HERE, '..', '..'))
PYTHON  = r'D:\Freelaces\Server\python.exe'
MAIN_PY = os.path.join(_HERE, '..', 'main.py')
IN_DIR  = os.path.join(_ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii')
OUT_DIR = os.path.join(IN_DIR, 'hasil')

ARGS = [
    'paket3', 'Times New Roman', '12 pt',
    'Ya',           # hidden_cov
    'Tengah Bawah', # posisi roman
    'Tengah Bawah', # pos_bab
    'Kanan Atas',   # pos_isi_bab
    'iii',          # dimulai_dari
    'Tidak',        # semb_dafus
    'Tidak',        # semb_lamprn
    '2',            # num_cover
]

os.makedirs(OUT_DIR, exist_ok=True)

files = sorted(f for f in os.listdir(IN_DIR) if f.lower().endswith('.docx'))
passed = failed = errored = 0

for fname in files:
    in_path  = os.path.join(IN_DIR,  fname)
    stem     = os.path.splitext(fname)[0]
    out_path = os.path.join(OUT_DIR, f"{stem}_c2.docx")
    cmd = [PYTHON, MAIN_PY, in_path, out_path] + ARGS
    res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    if res.returncode == 0:
        passed += 1
        print(f"  OK   {fname}  ->  {os.path.basename(out_path)}")
    else:
        errored += 1
        msg = (res.stderr or res.stdout or '').strip()[:120]
        print(f"  ERR  {fname}  {msg}")

print(f"\n{'='*50}")
print(f"  {passed}/{passed+errored} OK  |  {errored} ERR")
print(f"  Output: {OUT_DIR}")
print(f"{'='*50}")
