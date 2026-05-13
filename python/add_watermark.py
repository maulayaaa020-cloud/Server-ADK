"""
add_watermark.py — Tambah watermark diagonal ke setiap halaman PDF.
Usage: python add_watermark.py <input.pdf> <output.pdf> [teks watermark]
"""
import sys
import json
import os
from io import BytesIO

WATERMARK_TEXT = "PREVIEW HASIL ADK"


def buat_halaman_watermark(width, height, text):
    """
    Tile watermark menggunakan koordinat (along, perp):
    - along: searah teks (cos θ, sin θ)
    - perp : tegak lurus teks (−sin θ, cos θ)
    Pendekatan ini menjamin semua 4 sudut halaman tertutup.
    """
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import Color
    import math

    ANGLE     = 35          # derajat
    TEXT_STEP = 230         # jarak antar teks searah diagonal
    LINE_STEP = 95          # jarak antar baris (tegak lurus diagonal)
    FONT_SIZE = 33          # ukuran font asli

    angle_rad = math.radians(ANGLE)
    dx =  math.cos(angle_rad)   # arah teks
    dy =  math.sin(angle_rad)
    nx = -math.sin(angle_rad)   # tegak lurus teks
    ny =  math.cos(angle_rad)

    # Koordinat (along, perp) dari keempat sudut halaman
    corners = [(0, 0), (width, 0), (0, height), (width, height)]
    along_vals = [x * dx + y * dy for x, y in corners]
    perp_vals  = [x * nx + y * ny for x, y in corners]

    along_min = min(along_vals) - TEXT_STEP
    along_max = max(along_vals) + TEXT_STEP
    perp_min  = min(perp_vals)  - LINE_STEP
    perp_max  = max(perp_vals)  + LINE_STEP

    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    c.setFont("Helvetica-Bold", FONT_SIZE)
    c.setFillColor(Color(0.50, 0.50, 0.50, alpha=0.20))

    line = 0
    perp = perp_min
    while perp <= perp_max:
        # Geser selang-seling agar pola berlian, bukan kotak
        along = along_min + (line % 2) * (TEXT_STEP / 2)
        while along <= along_max:
            # Ubah (along, perp) → (x, y) di halaman
            x = along * dx + perp * nx
            y = along * dy + perp * ny

            c.saveState()
            c.translate(x, y)
            c.rotate(ANGLE)
            c.drawCentredString(0, 0, text)
            c.restoreState()

            along += TEXT_STEP
        perp += LINE_STEP
        line += 1

    c.save()
    packet.seek(0)
    return packet


def tambah_watermark(input_pdf, output_pdf, text=WATERMARK_TEXT):
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    # append() meng-clone seluruh struktur PDF dengan benar ke writer
    writer.append(input_pdf)

    # Cache watermark per ukuran halaman agar tidak dibuat ulang terus
    wm_cache = {}

    for page in writer.pages:
        w = round(float(page.mediabox.width),  1)
        h = round(float(page.mediabox.height), 1)
        key = (w, h)

        if key not in wm_cache:
            wm_packet = buat_halaman_watermark(w, h, text)
            wm_reader  = PdfReader(wm_packet)
            wm_cache[key] = (wm_reader, wm_reader.pages[0])

        _, wm_page = wm_cache[key]
        page.merge_page(wm_page, over=True)

    with open(output_pdf, "wb") as f:
        writer.write(f)


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"ok": False, "error": "Usage: add_watermark.py <input.pdf> <output.pdf> [text]"}))
        sys.exit(1)

    input_pdf  = sys.argv[1]
    output_pdf = sys.argv[2]
    text       = sys.argv[3] if len(sys.argv) > 3 else WATERMARK_TEXT

    if not os.path.exists(input_pdf):
        print(json.dumps({"ok": False, "error": f"File tidak ditemukan: {input_pdf}"}))
        sys.exit(1)

    try:
        tambah_watermark(input_pdf, output_pdf, text)
        if os.path.exists(output_pdf):
            print(json.dumps({"ok": True}))
        else:
            print(json.dumps({"ok": False, "error": "Output PDF tidak terbentuk"}))
            sys.exit(1)
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
