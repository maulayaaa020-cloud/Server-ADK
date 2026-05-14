"""
convert_pdf.py — Konversi file .docx ke .pdf.
Linux: menggunakan LibreOffice (--headless)
Windows: menggunakan docx2pdf (Microsoft Word)
Usage: python convert_pdf.py <input.docx> <output.pdf>
"""
import sys
import json
import os
import platform


def convert_with_libreoffice(input_docx, output_pdf):
    import subprocess
    import shutil

    out_dir = os.path.dirname(os.path.abspath(output_pdf))
    if not out_dir:
        out_dir = '.'

    env = {"HOME": "/tmp", "PATH": "/usr/bin:/usr/local/bin:/bin:/usr/sbin:/sbin"}
    result = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', out_dir, input_docx],
        capture_output=True, text=True, timeout=120, env=env
    )

    # LibreOffice outputs basename_of_input.pdf in out_dir
    base = os.path.splitext(os.path.basename(input_docx))[0]
    lo_output = os.path.join(out_dir, base + '.pdf')

    if os.path.abspath(lo_output) != os.path.abspath(output_pdf) and os.path.exists(lo_output):
        shutil.move(lo_output, output_pdf)

    if not os.path.exists(output_pdf):
        raise Exception(f"LibreOffice gagal: {result.stderr or result.stdout}")


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"ok": False, "error": "Usage: python convert_pdf.py <input.docx> <output.pdf>"}))
        sys.exit(1)

    input_docx = sys.argv[1]
    output_pdf  = sys.argv[2]

    if not os.path.exists(input_docx):
        print(json.dumps({"ok": False, "error": f"File tidak ditemukan: {input_docx}"}))
        sys.exit(1)

    try:
        if platform.system() == 'Windows':
            from docx2pdf import convert
            convert(input_docx, output_pdf)
        else:
            convert_with_libreoffice(input_docx, output_pdf)

        if os.path.exists(output_pdf):
            print(json.dumps({"ok": True}))
        else:
            print(json.dumps({"ok": False, "error": "PDF tidak terbentuk setelah konversi"}))
            sys.exit(1)

    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
