"""
convert_pdf.py — Konversi file .docx ke .pdf menggunakan Microsoft Word (via docx2pdf).
Usage: python convert_pdf.py <input.docx> <output.pdf>
"""
import sys
import json
import os


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
        from docx2pdf import convert
        convert(input_docx, output_pdf)

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
