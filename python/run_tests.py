"""
run_tests.py — Regression test suite untuk ADK page numbering.

Usage:
  python run_tests.py              Jalankan semua test, bandingkan dengan expected
  python run_tests.py --lock       Simpan output saat ini sebagai expected (setelah konfirmasi manual)
  python run_tests.py --add <path> Tambah file baru ke test suite (interaktif)
  python run_tests.py --show       Tampilkan semua expected results

Cara kerja:
  1. Setelah konfirmasi manual bahwa N file sudah benar → jalankan --lock
  2. Saat akan upgrade/fix file baru → jalankan tanpa argumen (run all)
  3. Jika semua PASS berarti tidak ada regresi, aman untuk push
  4. Setelah file baru benar secara manual → jalankan --add <path>
"""
import sys, json, os, subprocess

_HERE   = os.path.dirname(os.path.abspath(__file__))
_ROOT   = os.path.normpath(os.path.join(_HERE, '..'))
CASES_F = os.path.join(_HERE, 'test_cases.json')
MAIN_PY = os.path.join(_HERE, 'main.py')
PYTHON  = r'D:\Freelaces\Server\python.exe'
TMP_OUT = os.path.join(_ROOT, 'output', '_test_tmp.docx')

# Parameter default untuk semua test — harus sama persis dengan form website:
#   paket3 | Times New Roman | 12pt | sembunyikan cover | Tengah Bawah
#   pos_bab=Tengah Bawah | pos_isi_bab=Kanan Atas | dimulai=ii | 1 cover
DEFAULT_ARGS = [
    'paket3', 'Times New Roman', '12 pt',
    'Ya',           # hidden_cov
    'Tengah Bawah', # posisi roman
    'Tengah Bawah', # pos_bab  (paket4 only, ignored paket3)
    'Kanan Atas',   # pos_isi_bab (paket4 only, ignored paket3)
    'ii',           # dimulai_dari
    'Tidak',        # semb_dafus
    'Tidak',        # semb_lamprn
    '1',            # num_cover
]


# ── helpers ───────────────────────────────────────────────────────────────────

def run_file(input_path, args=None):
    cmd = [PYTHON, MAIN_PY, input_path, TMP_OUT] + (args or DEFAULT_ARGS)
    res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    try:
        return json.loads(res.stdout)
    except Exception:
        return {'status': 'error', 'message': (res.stderr or res.stdout or '').strip()[:200]}


def load_cases():
    if os.path.exists(CASES_F):
        with open(CASES_F, 'r', encoding='utf-8-sig') as f:
            return json.load(f)
    return {}


def save_cases(cases):
    with open(CASES_F, 'w', encoding='utf-8') as f:
        json.dump(cases, f, ensure_ascii=False, indent=2)


def read_page_nums(docx_path):
    """Baca pgNumType (fmt + start) dari setiap section di output .docx.
    Ini yang membuktikan dimulai_dari=ii benar-benar diterapkan ke file."""
    try:
        sys.path.insert(0, _HERE)
        from docx import Document
        from docx.oxml.ns import qn
        doc = Document(docx_path)
        result = []
        for sec in doc.sections:
            pn    = sec._sectPr.find(qn('w:pgNumType'))
            fmt   = pn.get(qn('w:fmt'),   'decimal') if pn is not None else 'decimal'
            start = pn.get(qn('w:start'), None)      if pn is not None else None
            result.append({'fmt': fmt, 'start': start})
        return result
    except Exception:
        return []


def snapshot(result):
    """Ambil bagian yang relevan dari output main.py + page numbers dari output file."""
    snap = {
        'detected_bab':   result.get('detected_bab', []),
        'total_sections': result.get('total_sections', 0),
        'sections':       result.get('sections', []),
    }
    if os.path.exists(TMP_OUT):
        snap['page_nums'] = read_page_nums(TMP_OUT)
    return snap


def diff_results(expected, actual):
    """Bandingkan expected vs actual. Kembalikan list string perbedaan."""
    diffs = []
    if expected.get('detected_bab') != actual.get('detected_bab'):
        diffs.append(
            f"  detected_bab BERUBAH:\n"
            f"    exp: {expected['detected_bab']}\n"
            f"    got: {actual['detected_bab']}"
        )
    if expected.get('total_sections') != actual.get('total_sections'):
        diffs.append(
            f"  total_sections: exp={expected['total_sections']} got={actual['total_sections']}"
        )
    exp_secs = expected.get('sections', [])
    got_secs = actual.get('sections', [])
    for i, (es, gs) in enumerate(zip(exp_secs, got_secs)):
        if es.get('first_content') != gs.get('first_content'):
            diffs.append(
                f"  section[{i}] first_content:\n"
                f"    exp: {repr(es.get('first_content'))}\n"
                f"    got: {repr(gs.get('first_content'))}"
            )
    # Cek page numbers (membuktikan dimulai_dari diterapkan ke file output)
    exp_pn = expected.get('page_nums', [])
    got_pn = actual.get('page_nums', [])
    if exp_pn and got_pn:
        for i, (ep, gp) in enumerate(zip(exp_pn, got_pn)):
            if ep.get('fmt') != gp.get('fmt') or ep.get('start') != gp.get('start'):
                diffs.append(
                    f"  section[{i}] page_num: exp={ep} got={gp}"
                )
    return diffs


# ── commands ──────────────────────────────────────────────────────────────────

def cmd_run(cases):
    """Jalankan semua test, bandingkan dengan expected."""
    if not cases:
        print("Belum ada test case. Jalankan --lock terlebih dahulu.")
        return
    passed = failed = errored = skipped = 0
    for name, meta in cases.items():
        path = meta.get('path', '')
        if not os.path.exists(path):
            print(f"  SKIP {name}  (file tidak ada: {path})")
            skipped += 1
            continue
        result = run_file(path, meta.get('args'))
        if result.get('status') != 'success':
            errored += 1
            print(f"  ERR  {name}  {result.get('message','')[:80]}")
            continue
        snap  = snapshot(result)
        diffs = diff_results(meta, snap)
        if diffs:
            failed += 1
            print(f"  FAIL {name}")
            for d in diffs:
                print(d)
        else:
            passed += 1
            print(f"  PASS {name}")
    total = passed + failed + errored
    print(f"\n{'='*50}")
    print(f"  {passed}/{total} PASS  |  {failed} FAIL  |  {errored} ERR  |  {skipped} SKIP")
    print(f"{'='*50}")
    if failed or errored:
        sys.exit(1)


def cmd_lock(cases):
    """Jalankan semua file yang sudah ada, simpan output sekarang sebagai expected."""
    if not cases:
        print("Tidak ada test case. Gunakan --add untuk menambah file dulu.")
        return
    print("Locking expected results dari output saat ini...\n")
    locked = 0
    for name, meta in list(cases.items()):
        path = meta.get('path', '')
        if not os.path.exists(path):
            print(f"  SKIP {name}  (file tidak ada)")
            continue
        result = run_file(path, meta.get('args'))
        if result.get('status') != 'success':
            print(f"  ERR  {name}  {result.get('message','')[:80]}")
            continue
        snap = snapshot(result)
        cases[name].update(snap)
        locked += 1
        print(f"  LOCK {name}  |  bab: {snap['detected_bab']}")
    save_cases(cases)
    print(f"\n{locked} file di-lock -> {CASES_F}")


def cmd_add(cases, path):
    """Tambah file baru ke test suite secara interaktif."""
    path = os.path.abspath(path)
    if not os.path.exists(path):
        print(f"File tidak ditemukan: {path}")
        sys.exit(1)
    name = os.path.basename(path)
    if name in cases:
        print(f"{name} sudah ada di test suite. Gunakan --lock untuk update.")
        return

    print(f"Memproses {name}...")
    result = run_file(path)
    if result.get('status') != 'success':
        print(f"Error: {result.get('message', '')}")
        sys.exit(1)

    snap = snapshot(result)
    print(f"\nHasil untuk {name}:")
    print(f"  detected_bab   : {snap['detected_bab']}")
    print(f"  total_sections : {snap['total_sections']}")
    print(f"  sections       :")
    for s in snap['sections']:
        print(f"    [{s['index']}] {repr(s['first_content'])}")

    print(f"\nTambahkan {name} ke test suite dengan hasil ini? [y/N] ", end='')
    ans = input().strip().lower()
    if ans == 'y':
        cases[name] = {'path': path, **snap}
        save_cases(cases)
        print(f"Ditambahkan ke test suite: {name}  (total: {len(cases)} file)")
    else:
        print("Dibatalkan.")


def cmd_show(cases):
    """Tampilkan semua expected results."""
    if not cases:
        print("Belum ada test case.")
        return
    print(f"{'='*60}")
    for name, meta in cases.items():
        exists = os.path.exists(meta.get('path', ''))
        status = 'OK' if exists else 'MISSING'
        print(f"[{status}] {name}")
        print(f"  bab    : {meta.get('detected_bab', [])}")
        print(f"  secs   : {meta.get('total_sections')}")
    print(f"{'='*60}")
    print(f"Total: {len(cases)} file")


# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    cases = load_cases()
    cmd   = sys.argv[1] if len(sys.argv) > 1 else ''

    if cmd == '--lock':
        cmd_lock(cases)
    elif cmd == '--add':
        if len(sys.argv) < 3:
            print("Usage: run_tests.py --add <path/to/file.docx>")
            sys.exit(1)
        cmd_add(cases, sys.argv[2])
    elif cmd == '--show':
        cmd_show(cases)
    elif cmd in ('', '--run'):
        cmd_run(cases)
    else:
        print(__doc__)