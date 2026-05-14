import sys, json, os, platform

def convert_with_libreoffice(input_docx, output_pdf):
    import subprocess, shutil
    out_dir = os.path.dirname(os.path.abspath(output_pdf)) or '.'
    result = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', out_dir, input_docx],
        capture_output=True, text=True, timeout=120
    )
    base = os.path.splitext(os.path.basename(input_docx))[0]
    lo_out = os.path.join(out_dir, base + '.pdf')
    if os.path.abspath(lo_out) != os.path.abspath(output_pdf) and os.path.exists(lo_out):
        shutil.move(lo_out, output_pdf)
    if not os.path.exists(output_pdf):
        raise Exception(f"LibreOffice gagal: {result.stderr or result.stdout}")

def main():
    if len(sys.argv) < 3:
        print(json.dumps({"ok": False, "error": "Usage: convert_pdf.py <in.docx> <out.pdf>"})); sys.exit(1)
    input_docx, output_pdf = sys.argv[1], sys.argv[2]
    if not os.path.exists(input_docx):
        print(json.dumps({"ok": False, "error": f"File tidak ditemukan: {input_docx}"})); sys.exit(1)
    try:
        if platform.system() == 'Windows':
            from docx2pdf import convert
            convert(input_docx, output_pdf)
        else:
            convert_with_libreoffice(input_docx, output_pdf)
        if os.path.exists(output_pdf):
            print(json.dumps({"ok": True}))
        else:
            print(json.dumps({"ok": False, "error": "PDF tidak terbentuk"})); sys.exit(1)
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)})); sys.exit(1)

if __name__ == "__main__":
    main()
