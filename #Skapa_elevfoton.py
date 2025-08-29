#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import csv
import sys
import time
import socket
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from PIL import Image
import io

# ===== MAPPNAMN =====
PHOTO_DIR = "Foton"
CSV_DIR = "CSV"
OUT_BADGES = "Namnskyltar"
OUT_CLASS = "Klassfoton"
TXT_DIR = "Elevdata"  
OUT_CLASS = "Klassfoton"   

# ===== KONFIG / VERSION =====
version = "1.35"

# ===== STARTTID =====
script_start_ts = time.time()

# ===== LOGGFIL =====
os.makedirs(OUT_CLASS, exist_ok=True)   # Skapar mappen om den inte finns
LOG_FILE = os.path.join(OUT_CLASS, "log.txt")  # Sparar loggen i Klassfoton
if os.path.exists(LOG_FILE):
    os.remove(LOG_FILE)


def _now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def write_log_line(level, msg, console=True):
    timestamp = _now_ts()
    tag = "[INFO   ] " if level.upper() == "INFO" else "[SAKNAS ] " if "Saknad bild" in msg else "[ERROR  ] "
    line = f"{tag}{timestamp} - {msg}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line)
    if console:
        print(line.rstrip())

def write_class_header(class_name):
    header = f"-- Bearbetar klass: {class_name} --"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(header + "\n")
    print(header)

# ===== FRÅGA ANVÄNDAREN =====
year = input("Ange år (t.ex. 25): ").strip()
if not (year.isdigit() and len(year) == 2):
    print("Fel: ange år som två siffror, t.ex. 25")
    sys.exit(1)

create_csv = input("Vill du skapa CSV-filer? (Y/N): ").strip().lower() == 'y'
create_badges = input("Vill du skapa PDF: Namnskyltar? (Y/N): ").strip().lower() == 'y'
create_classphotos = input("Vill du skapa PDF: Klassfoton? (Y/N): ").strip().lower() == 'y'

computer = socket.gethostname()

# ===== LOGGA STARTINFO =====
with open(LOG_FILE, "a", encoding="utf-8") as f:
    f.write("="*70 + "\n")
    f.write(f"Skript startat: {_now_ts()}\n")
    f.write(f"Version: {version}\n")
    f.write(f"År: {year}\n")
    f.write(f"Skapa CSV: {'Ja' if create_csv else 'Nej'}\n")
    f.write(f"Skapa Namnskyltar: {'Ja' if create_badges else 'Nej'}\n")
    f.write(f"Skapa Klassfoton: {'Ja' if create_classphotos else 'Nej'}\n")
    f.write("="*70 + "\n")

print("Startar bearbetning... (se log.txt för detaljer)")

# ===== HJÄLP: Optimize image -> ImageReader eller None =====
def get_optimized_image_reader(path, target_w_mm, target_h_mm, dpi=200, quality=85):
    try:
        px_w = int(target_w_mm / mm * dpi / 25.4)
        px_h = int(target_h_mm / mm * dpi / 25.4)
        with Image.open(path) as img:
            img = ImageOps_exif_transpose_safe(img)
            img.thumbnail((px_w, px_h), Image.LANCZOS)
            buff = io.BytesIO()
            img.convert("RGB").save(buff, format="JPEG", quality=quality, optimize=True)
            buff.seek(0)
            return ImageReader(buff)
    except Exception as e:
        write_log_line("ERROR", f"BILDOPTIMERING FEL '{path}': {e}")
        return None

def ImageOps_exif_transpose_safe(img):
    try:
        from PIL import ImageOps
        return ImageOps.exif_transpose(img)
    except Exception:
        return img

# ===== RIT-FUNKTION: saknad bild =====
def draw_missing_photo(c, x, y, w, h, text="Bild saknas"):
    c.setStrokeColor(colors.red)
    c.setLineWidth(3)
    c.rect(x, y, w, h, stroke=1, fill=0)
    c.setStrokeColor(colors.red)
    c.setLineWidth(2)
    c.line(x, y, x + w, y + h)
    c.line(x + w, y, x, y + h)
    font_name = "Helvetica-Bold"
    font_size = 14
    c.setFont(font_name, font_size)
    text_w = c.stringWidth(text, font_name, font_size)
    text_h = font_size
    padding = 4
    rect_w = text_w + padding * 2
    rect_h = text_h + padding
    rect_x = x + (w - rect_w) / 2
    rect_y = y + (h - rect_h) / 2
    c.setFillColor(colors.white)
    c.setStrokeColor(colors.white)
    c.rect(rect_x, rect_y, rect_w, rect_h, stroke=0, fill=1)
    c.setFillColor(colors.black)
    c.setFont(font_name, font_size)
    c.drawCentredString(x + w / 2, y + h / 2 - (text_h * 0.08), text)

# ===== LÄS OCH PARSA .txt-FILER =====
def parse_txt_file(txt_path, class_name, year):
    rows = []
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = [ln.strip() for ln in f.readlines() if ln.strip()]
    except Exception as e:
        write_log_line("ERROR", f"Kan inte läsa {txt_path}: {e}")
        return rows
    for ln in lines:
        parts = ln.split()
        if len(parts) == 0:
            continue
        if len(parts) >= 3:
            email = parts[-1]
            names = parts[:-1]
        else:
            email = ""
            names = parts
        if len(names) == 1:
            firstname = names[0]
            surname = ""
        elif len(names) == 2:
            firstname, surname = names
        else:
            firstname = names[0]
            surname = " ".join(names[1:])
        image_path = os.path.join(PHOTO_DIR, f"{class_name}_{surname}_{firstname}.jpg")
        rows.append({
            "firstname": firstname,
            "surname": surname,
            "@image": image_path,
            "year": year,
            "class": class_name,
            "email": email
        })
    return rows

# ===== SKAPA CSV =====
def write_csv(csv_path, rows):
    try:
        os.makedirs(os.path.dirname(csv_path), exist_ok=True)
        with open(csv_path, "w", newline="", encoding="cp1252") as csvf:
            writer = csv.writer(csvf, delimiter=';')
            writer.writerow(("year", "class", "surname", "firstname", "@image", "email"))
            for r in rows:
                writer.writerow((r["year"], r["class"], r["surname"], r["firstname"], r["@image"], r.get("email", "")))
        return True
    except Exception as e:
        write_log_line("ERROR", f"Fel vid skrivning av CSV '{csv_path}': {e}")
        return False

# ===== SKAPA NAMNSKYLTAR =====
def create_pdf_badges(pdf_path, rows, class_name):
    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
        page_w, page_h = landscape(A4)
        c.setTitle(f"{class_name} - Namnskyltar")
        max_width = page_w - 40*mm
        y_positions = [0.65, 0.5, 0.32]
        for r in rows:
            firstname = r["firstname"]
            surname = r["surname"]
            class_text = r["class"]
            fontsize_fn = 120
            c.setFont("Helvetica-Bold", fontsize_fn)
            while c.stringWidth(firstname, "Helvetica-Bold", fontsize_fn) > max_width and fontsize_fn > 10:
                fontsize_fn -= 1
                c.setFont("Helvetica-Bold", fontsize_fn)
            c.drawCentredString(page_w/2, page_h * y_positions[0], firstname)
            fontsize_sn = 96
            c.setFont("Helvetica-Bold", fontsize_sn)
            while c.stringWidth(surname, "Helvetica-Bold", fontsize_sn) > max_width and fontsize_sn > 10:
                fontsize_sn -= 1
                c.setFont("Helvetica-Bold", fontsize_sn)
            c.drawCentredString(page_w/2, page_h * y_positions[1], surname)
            fontsize_class = 80
            c.setFont("Helvetica-Bold", fontsize_class)
            while c.stringWidth(class_text, "Helvetica-Bold", fontsize_class) > max_width and fontsize_class > 10:
                fontsize_class -= 1
                c.setFont("Helvetica-Bold", fontsize_class)
            c.drawCentredString(page_w/2, page_h * y_positions[2], class_text)
            c.showPage()
        c.save()
        write_log_line("INFO", f"PDF Namnskyltar skapad: {os.path.normpath(pdf_path)}")
        return True
    except Exception as e:
        write_log_line("ERROR", f"Fel vid skapande av namnskyltar '{pdf_path}': {e}")
        return False

# ===== FOOTER FÖR KLASSFOTON =====
def draw_footer_with_mail(c, page_w, page_h, page_num, total_pages, version):
    timestamp = _now_ts()
    python_ver = sys.version.split()[0]
    reportlab_ver = getattr(sys.modules.get('reportlab'), 'Version', 'unknown')
    text_before = f"Skolfoton Realgymnasiet Gävle v {version} | Uppdaterad "
    text_after = f" med Python {python_ver}, ReportLab {reportlab_ver}. Skript av "
    author_text = "Andreas Hagström"
    mailto = "mailto:andreas.hagstrom@realgymnasiet.se"

    c.setFont("Helvetica", 8)
    total_text = text_before + timestamp + text_after + author_text + "."
    total_width = c.stringWidth(total_text, "Helvetica", 8)
    x_start = (page_w - total_width) / 2
    y_pos = 15*mm
    # Text before
    c.setFillColor(colors.black)
    c.drawString(x_start, y_pos, text_before)
    x_cursor = x_start + c.stringWidth(text_before, "Helvetica", 8)
    # Timestamp bold
    c.setFont("Helvetica-Bold", 8)
    c.drawString(x_cursor, y_pos, timestamp)
    x_cursor += c.stringWidth(timestamp, "Helvetica-Bold", 8)
    # Text after + author
    c.setFont("Helvetica", 8)
    c.drawString(x_cursor, y_pos, text_after + author_text + ".")
    author_width = c.stringWidth(author_text, "Helvetica", 8)
    c.linkURL(mailto, (x_cursor + c.stringWidth(text_after, "Helvetica", 8),
                        y_pos - 2,
                        x_cursor + c.stringWidth(text_after, "Helvetica", 8) + author_width,
                        y_pos + 8), relative=0, thickness=0, color=colors.blue)
    # Sidnummer
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8)
    c.drawCentredString(page_w/2, 10*mm, f"Sida {page_num} av {total_pages}")

# ===== SKAPA KLASSFOTON =====

def create_pdf_classphotos(pdf_path, rows, class_name, version):
    missing = 0
    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
        page_w, page_h = landscape(A4)
        c.setTitle(f"{class_name} - Klassfoton")
        cols = 4
        rows_per_page = 2
        top_margin = 30*mm
        bottom_margin = 20*mm
        left_margin = 25*mm
        right_margin = 25*mm
        usable_height = page_h - top_margin - bottom_margin
        usable_width = page_w - left_margin - right_margin
        cell_w = usable_width / cols
        cell_h = usable_height / rows_per_page
        photo_w = min(50*mm, cell_w * 0.9)
        photo_h = min(60*mm, cell_h * 0.75)
        total_pages = (len(rows) + (cols*rows_per_page - 1)) // (cols*rows_per_page)

        def draw_header():
            # Klassnamn
            c.setFont("Helvetica-Bold", 36)
            c.drawCentredString(page_w/2, page_h - 20*mm, class_name)

            # Antal elever
            student_count_text = f"Klassen innehåller {len(rows)} elever"
            c.setFont("Helvetica", 12)
            c.drawCentredString(page_w/2, page_h - 27*mm, student_count_text)

        count = 0
        page_num = 1
        draw_header()
        for idx, r in enumerate(rows):
            if idx > 0 and idx % (cols * rows_per_page) == 0:
                draw_footer_with_mail(c, page_w, page_h, page_num, total_pages, version)
                c.showPage()
                page_num += 1
                draw_header()
            page_index_in_page = idx % (cols * rows_per_page)
            row_idx = page_index_in_page // cols
            col_idx = page_index_in_page % cols
            x = left_margin + col_idx * cell_w + (cell_w - photo_w)/2
            y = page_h - top_margin - row_idx * cell_h - photo_h - ((cell_h - photo_h)/2)
            firstname = r["firstname"]
            surname = r["surname"]
            image_path = r["@image"]
            img_reader = None
            if image_path and os.path.exists(image_path):
                img_reader = get_optimized_image_reader(image_path, photo_w, photo_h)
            if img_reader:
                try:
                    c.drawImage(img_reader, x, y, width=photo_w, height=photo_h, preserveAspectRatio=True, anchor='c')
                except Exception as e:
                    draw_missing_photo(c, x, y, photo_w, photo_h)
                    missing += 1
                    write_log_line("ERROR", f"Saknad bild för {firstname} {surname}: {os.path.normpath(image_path)}")
            else:
                draw_missing_photo(c, x, y, photo_w, photo_h)
                missing += 1
                write_log_line("ERROR", f"Saknad bild för {firstname} {surname}: {os.path.normpath(image_path)}")
            c.setFont("Helvetica", 12)
            c.setFillColor(colors.black)
            c.drawCentredString(x + photo_w/2, y - (6*mm), f"{firstname} {surname}")
        draw_footer_with_mail(c, page_w, page_h, page_num, total_pages, version)
        c.save()
        write_log_line("INFO", f"PDF Klassfoton skapad: {os.path.normpath(pdf_path)} (saknade bilder: {missing})")
        return missing
    except Exception as e:
        write_log_line("ERROR", f"Fel vid skapande av klassfoton för {class_name}: {e}")
        return missing


# ===== HUVUDLOOP =====
txt_files = sorted([f for f in os.listdir(TXT_DIR) if f.lower().endswith('.txt') and not f.lower().startswith('log')])
if not txt_files:
    write_log_line("ERROR", f"Inga .txt-filer hittades i mappen '{TXT_DIR}'. Avslutar.")
    sys.exit(1)

total_missing = 0
for txt_file in txt_files:
    class_name = os.path.splitext(os.path.basename(txt_file))[0]
    write_class_header(class_name)
    txt_file_path = os.path.join(TXT_DIR, txt_file)
    parsed_rows = parse_txt_file(txt_file_path, class_name, year)

    if not parsed_rows:
        write_log_line("ERROR", f"Inga elever i {txt_file} eller läsfel.")
        continue

    csv_path = os.path.join(CSV_DIR, f"{class_name}.csv")
    if create_csv:
        ok = write_csv(csv_path, parsed_rows)
        if ok:
            write_log_line("INFO", f"CSV skapad: {os.path.normpath(csv_path)}")

    if create_badges:
        pdf_badges_path = os.path.join(OUT_BADGES, f"{class_name}.pdf")
        create_pdf_badges(pdf_badges_path, parsed_rows, class_name)

    if create_classphotos:
        pdf_class_path = os.path.join(OUT_CLASS, f"{class_name}.pdf")
        missing = create_pdf_classphotos(pdf_class_path, parsed_rows, class_name, version)
        total_missing += missing

    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write("\n")
    print("")

# ===== AVSLUTA OCH LOGGA SAMMANFATTNING =====
script_end_ts = time.time()
duration = script_end_ts - script_start_ts
with open(LOG_FILE, "a", encoding="utf-8") as f:
    f.write("="*70 + "\n")
    f.write(f"Stopptid: {_now_ts()}\n")
    f.write(f"Total tid: {duration:.2f} sekunder\n")
    f.write(f"Totalt saknade bilder (alla klasser): {total_missing}\n")
    f.write(f"PC: {computer}\n")
    f.write("Bearbetning klar\n")

print("="*70)
print(f"Stopptid: {_now_ts()}")
print(f"Total tid: {duration:.2f} sekunder")
print(f"Totalt saknade bilder (alla klasser): {total_missing}")
print(f"PC: {computer}")
print("Bearbetning klar")
