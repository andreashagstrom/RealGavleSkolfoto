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
from PIL import Image, ImageOps
import io
import subprocess

# ===== 1. DEPENDENCIES-CHECK =====
def check_dependencies():
    dependencies = ["PIL", "reportlab"]
    for dep in dependencies:
        try:
            __import__(dep)
            print(f"[INFO] Dependency '{dep}' är installerad.")
        except ImportError:
            ans = input(f"[VARNING] Dependency '{dep}' saknas. Vill du installera den nu? (Y/N): ").strip().lower()
            if ans == 'y':
                subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
                print(f"[INFO] '{dep}' installerad.")
            else:
                print(f"[ERROR] '{dep}' saknas och behövs för att skriptet ska fungera. Avslutar.")
                sys.exit(1)

check_dependencies()

# ===== 2. KONFIG / VERSION =====
version = "1.39"

# ===== STARTTEXT I TERMINAL =====
print(f"\nSkolfoton Realgymnasiet Gävle v {version}\n")

# ===== 3. VILKA MAPPAR SKALL KÖRAS? ===== #
base_folders_input = input("Ange en eller flera mappnamn (t.ex. 23,24,25): ").strip()
if not base_folders_input:
    print("Fel: du måste ange minst en mapp")
    sys.exit(1)

# Gör om till lista och trimma bort mellanslag
base_folders = [m.strip() for m in base_folders_input.split(",") if m.strip()]
if not base_folders:
    print("Fel: inga giltiga mappar angavs")
    sys.exit(1)

# ===== 4. STARTTID =====
script_start_ts = time.time()
computer = socket.gethostname()

def _now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def write_log_line(log_file, level, msg, console=True):
    timestamp = _now_ts()
    tag = "[INFO   ] " if level.upper() == "INFO" else "[SAKNAS ] " if "Saknad bild" in msg else "[ERROR  ] "
    line = f"{tag}{timestamp} - {msg}\n"
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(line)
    if console:
        print(line.rstrip())

def write_class_header(log_file, class_name):
    header = f"-- Bearbetar klass: {class_name} --"
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(header + "\n")
    print(header)

# ===== 5. ANVÄNDARVAL =====
create_csv = input("Vill du skapa CSV-filer? (Y/N): ").strip().lower() == 'y'
create_badges = input("Vill du skapa PDF: Namnskyltar? (Y/N): ").strip().lower() == 'y'
create_classphotos = input("Vill du skapa PDF: Klassfoton? (Y/N): ").strip().lower() == 'y'

print("Startar bearbetning... (se respektive log.txt för detaljer)")

# ===== 6. ANVÄNDARVAL: BILDKVALITET =====
def choose_image_quality():
    while True:
        user_quality = input("Ange önskad bildkvalitet (1-100): ").strip()
        if user_quality.isdigit() and 1 <= int(user_quality) <= 100:
            user_quality = int(user_quality)
            print(f"Vald bildkvalitet: {user_quality}")
            return user_quality
        else:
            print("Felaktig inmatning. Ange en siffra mellan 1-100.")

image_quality = choose_image_quality()

# ===== 7. OPTIMERA BILD =====
def get_optimized_image_reader(path, target_w_mm, target_h_mm, dpi=200, quality=85):
    try:
        px_w = int(target_w_mm / mm * dpi / 25.4)
        px_h = int(target_h_mm / mm * dpi / 25.4)
        with Image.open(path) as img:
            img = ImageOps.exif_transpose(img)
            img.thumbnail((px_w, px_h), Image.LANCZOS)
            buff = io.BytesIO()
            img.convert("RGB").save(buff, format="JPEG", quality=quality, optimize=True)
            buff.seek(0)
            return ImageReader(buff)
    except Exception as e:
        return None

# ===== 8. RIT-FUNKTION: SAKNAD BILD =====
def draw_missing_photo(c, x, y, w, h, text="Bild saknas"):
    c.setStrokeColor(colors.red)
    c.setLineWidth(3)
    c.rect(x, y, w, h, stroke=1, fill=0)
    c.setLineWidth(2)
    c.line(x, y, x + w, y + h)
    c.line(x + w, y, x, y + h)
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.black)
    c.drawCentredString(x + w/2, y + h/2 - 7, text)

# ===== 9. LÄS OCH PARSA .txt-FILER =====
def parse_txt_file(txt_path, class_name, year, photo_dir):
    rows = []
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = [ln.strip() for ln in f.readlines() if ln.strip()]
    except Exception:
        return rows
    for ln in lines:
        parts = ln.split()
        if len(parts) >= 3:
            email = parts[-1]
            names = parts[:-1]
        else:
            email = ""
            names = parts
        firstname = names[0]
        surname = " ".join(names[1:]) if len(names) > 1 else ""
        image_path = os.path.join(photo_dir, f"{class_name}_{surname}_{firstname}.jpg")
        rows.append({
            "firstname": firstname,
            "surname": surname,
            "@image": image_path,
            "year": year,
            "class": class_name,
            "email": email
        })
    return rows

# ===== 10. SKAPA CSV =====
def write_csv(csv_path, rows, log_file):
    try:
        os.makedirs(os.path.dirname(csv_path), exist_ok=True)
        with open(csv_path, "w", newline="", encoding="cp1252") as csvf:
            writer = csv.writer(csvf, delimiter=';')
            writer.writerow(("year", "class", "surname", "firstname", "@image", "email"))
            for r in rows:
                writer.writerow((r["year"], r["class"], r["surname"], r["firstname"], r["@image"], r.get("email", "")))
        write_log_line(log_file, "INFO", f"CSV skapad: {os.path.normpath(csv_path)}")
    except Exception as e:
        write_log_line(log_file, "ERROR", f"Fel vid skrivning av CSV '{csv_path}': {e}")

# ===== 11. SKAPA NAMNSKYLTAR =====
def create_pdf_badges(pdf_path, rows, class_name, log_file):
    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
        page_w, page_h = landscape(A4)
        c.setTitle(f"{class_name} - Namnskyltar")
        max_width = page_w - 40*mm
        y_positions = [0.65, 0.5, 0.32]
        for r in rows:
            firstname, surname, class_text = r["firstname"], r["surname"], r["class"]
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
        write_log_line(log_file, "INFO", f"PDF Namnskyltar skapad: {os.path.normpath(pdf_path)}")
    except Exception as e:
        write_log_line(log_file, "ERROR", f"Fel vid skapande av namnskyltar '{pdf_path}': {e}")

# ===== 12. FOOTER =====
def draw_footer_with_mail(c, page_w, page_h, page_num, total_pages, version):
    c.setFont("Helvetica", 8)
    c.drawCentredString(page_w/2, 10*mm, f"Sida {page_num} av {total_pages}")

# ===== 13. SKAPA KLASSFOTON =====
def create_pdf_classphotos(pdf_path, rows, class_name, version, quality, log_file, missing_list):
    missing = 0
    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
        page_w, page_h = landscape(A4)
        c.setTitle(f"{class_name} - Klassfoton")
        cols, rows_per_page = 4, 2
        top_margin, bottom_margin = 30*mm, 20*mm
        left_margin, right_margin = 25*mm, 25*mm
        usable_height = page_h - top_margin - bottom_margin
        usable_width = page_w - left_margin - right_margin
        cell_w, cell_h = usable_width / cols, usable_height / rows_per_page
        photo_w, photo_h = min(50*mm, cell_w * 0.9), min(60*mm, cell_h * 0.75)
        total_pages = (len(rows) + (cols*rows_per_page - 1)) // (cols*rows_per_page)

        def draw_header():
            c.setFont("Helvetica-Bold", 36)
            c.drawCentredString(page_w/2, page_h - 20*mm, class_name)
            student_count_text = f"Klassen har {len(rows)} elever."
            c.setFont("Helvetica", 12)
            c.drawCentredString(page_w/2, page_h - 27*mm, student_count_text)

        page_num = 1
        draw_header()
        for idx, r in enumerate(rows):
            if idx > 0 and idx % (cols * rows_per_page) == 0:
                draw_footer_with_mail(c, page_w, page_h, page_num, total_pages, version)
                c.showPage()
                page_num += 1
                draw_header()
            row_idx, col_idx = (idx % (cols * rows_per_page)) // cols, (idx % (cols * rows_per_page)) % cols
            x = left_margin + col_idx * cell_w + (cell_w - photo_w)/2
            y = page_h - top_margin - row_idx * cell_h - photo_h - ((cell_h - photo_h)/2)
            firstname, surname, image_path = r["firstname"], r["surname"], r["@image"]
            img_reader = None
            if image_path and os.path.exists(image_path):
                img_reader = get_optimized_image_reader(image_path, photo_w, photo_h, quality=quality)
            if img_reader:
                try:
                    c.drawImage(img_reader, x, y, width=photo_w, height=photo_h, preserveAspectRatio=True, anchor='c')
                except Exception:
                    draw_missing_photo(c, x, y, photo_w, photo_h)
                    missing += 1
                    msg = f"Saknad bild för {firstname} {surname}: {os.path.normpath(image_path)}"
                    write_log_line(log_file, "SAKNAS", msg)
                    missing_list.append(r)
            else:
                draw_missing_photo(c, x, y, photo_w, photo_h)
                missing += 1
                msg = f"Saknad bild för {firstname} {surname}: {os.path.normpath(image_path)}"
                write_log_line(log_file, "SAKNAS", msg)
                missing_list.append(r)
            c.setFont("Helvetica", 12)
            c.drawCentredString(x + photo_w/2, y - (6*mm), f"{firstname} {surname}")
        draw_footer_with_mail(c, page_w, page_h, page_num, total_pages, version)
        c.save()
        write_log_line(log_file, "INFO", f"PDF Klassfoton skapad: {os.path.normpath(pdf_path)} (saknade bilder: {missing})")
        return missing
    except Exception as e:
        write_log_line(log_file, "ERROR", f"Fel vid skapande av klassfoton för {class_name}: {e}")
        return missing

# ===== 14. HUVUDLOOP FÖR VARJE MAPP =====
total_missing_global = 0
for base_folder in base_folders:
    PHOTO_DIR = os.path.join(base_folder, "Foton")
    CSV_DIR = os.path.join(base_folder, "CSV")
    OUT_BADGES = os.path.join(base_folder, "Namnskyltar")
    OUT_CLASS = os.path.join(base_folder, "Klassfoton")
    TXT_DIR = os.path.join(base_folder, "Elevdata")
    os.makedirs(OUT_CLASS, exist_ok=True)
    LOG_FILE = os.path.join(OUT_CLASS, "log.txt")
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)

    year = base_folder
    all_missing_students = []

    txt_files = sorted([f for f in os.listdir(TXT_DIR) if f.lower().endswith('.txt')])
    if not txt_files:
        write_log_line(LOG_FILE, "ERROR", f"Inga .txt-filer hittades i mappen '{TXT_DIR}'. Avslutar.")
        continue

    total_missing = 0
    for txt_file in txt_files:
        class_name = os.path.splitext(os.path.basename(txt_file))[0]
        write_class_header(LOG_FILE, class_name)
        txt_file_path = os.path.join(TXT_DIR, txt_file)
        parsed_rows = parse_txt_file(txt_file_path, class_name, year, PHOTO_DIR)

        if not parsed_rows:
            write_log_line(LOG_FILE, "ERROR", f"Inga elever i {txt_file} eller läsfel.")
            continue

        if create_csv:
            csv_path = os.path.join(CSV_DIR, f"{class_name}.csv")
            write_csv(csv_path, parsed_rows, LOG_FILE)

        if create_badges:
            pdf_badges_path = os.path.join(OUT_BADGES, f"{class_name}.pdf")
            create_pdf_badges(pdf_badges_path, parsed_rows, class_name, LOG_FILE)

        if create_classphotos:
            pdf_class_path = os.path.join(OUT_CLASS, f"{class_name}.pdf")
            missing = create_pdf_classphotos(pdf_class_path, parsed_rows, class_name, version, image_quality, LOG_FILE, all_missing_students)
            total_missing += missing

    if all_missing_students:
        os.makedirs(OUT_BADGES, exist_ok=True)
        missing_pdf_path = os.path.join(OUT_BADGES, "EleverSomSaknas.pdf")
        create_pdf_badges(missing_pdf_path, all_missing_students, "Elever som saknas", LOG_FILE)
        write_log_line(LOG_FILE, "INFO", f"PDF med saknade elever skapad: {os.path.normpath(missing_pdf_path)}")

    total_missing_global += total_missing

# ===== 15. SLUTRAPPORT =====
script_end_ts = time.time()
duration = script_end_ts - script_start_ts

print("="*70)
print(f"Stopptid: {_now_ts()}")
print(f"Total tid: {duration:.2f} sekunder")
print(f"Totalt saknade bilder (alla mappar): {total_missing_global}")
print(f"PC: {computer}")
print("Bearbetning klar")
