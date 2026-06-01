"""
run_preview_folders.py — Jalankan semua file di dua folder test dan simpan
output ke subfolder hasil/ masing-masing.

  cover 1 dimulai dari ii  : paket3, num_cover=1, dimulai_dari=ii
  cover 2 dimulai dari iii : paket3, num_cover=2, dimulai_dari=iii
"""
import sys, os, subprocess

sys.stdout.reconfigure(encoding='utf-8')

_HERE    = os.path.dirname(os.path.abspath(__file__))
_ROOT    = os.path.normpath(os.path.join(_HERE, '..', '..'))
MAIN_PY  = os.path.join(_ROOT, 'python', 'main.py')
PYTHON   = r'D:\Freelaces\Server\python.exe'

FOLDERS = [
    {
        'src': os.path.join(_ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii'),
        'out': os.path.join(_ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'hasil'),
        'args': ['paket3', 'Times New Roman', '12 pt',
                 'Ya', 'Tengah Bawah', 'Tengah Bawah', 'Kanan Atas',
                 'ii', 'Tidak', 'Tidak', '1'],
        'label': 'cover 1 dimulai dari ii',
    },
    {
        'src': os.path.join(_ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii'),
        'out': os.path.join(_ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'hasil'),
        'args': ['paket3', 'Times New Roman', '12 pt',
                 'Ya', 'Tengah Bawah', 'Tengah Bawah', 'Kanan Atas',
                 'iii', 'Tidak', 'Tidak', '2'],
        'label': 'cover 2 dimulai dari iii',
    },
]


def run_folder(cfg):
    src_dir = cfg['src']
    out_dir = cfg['out']
    args    = cfg['args']
    label   = cfg['label']

    os.makedirs(out_dir, exist_ok=True)
    docx_files = sorted(
        f for f in os.listdir(src_dir)
        if f.lower().endswith('.docx') and not f.startswith('~')
    )

    passed = failed = errored = 0
    print(f'\n{"="*60}')
    print(f'  {label}  ({len(docx_files)} file)')
    print(f'  Output → {out_dir}')
    print(f'{"="*60}')

    for fname in docx_files:
        inp = os.path.join(src_dir, fname)
        out = os.path.join(out_dir, fname)
        cmd = [PYTHON, MAIN_PY, inp, out] + args
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        try:
            import json as _json
            data = _json.loads(res.stdout)
            status = data.get('status', '')
        except Exception:
            status = 'error'
            data   = {'message': (res.stderr or res.stdout or '').strip()[:200]}

        if status == 'success':
            passed += 1
            bab = data.get('detected_bab', [])
            secs = data.get('total_sections', '?')
            print(f'  OK   {fname}  |  secs={secs}  bab={bab}')
        else:
            errored += 1
            msg = data.get('message', '')[:100]
            print(f'  ERR  {fname}  |  {msg}')

    print(f'\n  {passed} OK  |  {errored} ERR  (total {len(docx_files)})')
    return passed, errored


if __name__ == '__main__':
    total_ok = total_err = 0
    for cfg in FOLDERS:
        ok, err = run_folder(cfg)
        total_ok  += ok
        total_err += err

    print(f'\n{"="*60}')
    print(f'  TOTAL  {total_ok} OK  |  {total_err} ERR')
    print(f'{"="*60}')
