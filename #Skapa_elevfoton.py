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

# ===== 2. VERSION =====
version = "1.60"

print(f"\nSkolfoton Realgymnasiet Gävle v {version}\n")

# ===== 3. VÄLJ MAPPA/MAPPAR =====
base_folders_input = input("Ange en eller flera mappnamn (t.ex. 23,24,25): ").strip()
if not base_folders_input:
    print("Fel: du måste ange minst en mapp")
    sys.exit(1)

base_folders = [m.strip() for m in base_folders_input.split(",") if m.strip()]
if not base_folders:
    print("Fel: inga giltiga mappar angavs")
    sys.exit(1)

# ===== 4. STARTDATA =====
script_start_ts = time.time()
computer = socket.gethostname()

def _now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ===== 5. LOGGNING =====
def write_log_line(log_file, level, msg, console=True):
    timestamp = _now_ts()
    tag = "[INFO   ] " if level.upper() == "INFO" else "[SAKNAS ] " if "Saknad bild" in msg else "[ERROR  ] "
    line = f"{tag}{timestamp} - {msg}\n"
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(line)
    if console:
        print(line.rstrip())

def write_class_header(log_file, class_name):
    header = f"-- Bearbetar klass: {class_name} --"
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(header + "\n")
    print(header)

# ===== 6. ANVÄNDARVAL =====
create_csv = input("Vill du skapa CSV-filer? (Y/N): ").strip().lower() == 'y'
create_badges = input("Vill du skapa PDF: Namnskyltar? (Y/N): ").strip().lower() == 'y'
create_classphotos = input("Vill du skapa PDF: Klassfoton? (Y/N): ").strip().lower() == 'y'

print("Startar bearbetning... (se loggfil_YYYYMMDD.txt i respektive årsmapp för detaljer)")

# ===== 7. VÄLJ BILDKVALITET =====
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

# ===== 8. OPTIMERA BILD =====
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
    except Exception:
        return None

# ===== 9. PARSA TXT-FILER =====
def parse_txt_file(txt_path, class_name, year, photo_dir):
    rows = []
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = [ln.strip() for ln in f.readlines() if ln.strip()]
    except Exception:
        return rows

    for ln in lines:
        parts = ln.split()
        if len(parts) < 4:
            continue

        # Hämta fullständigt namn (mellan klasskod och sista fältet)
        fullname_parts = parts[1:-1]
        if len(fullname_parts) > 1:
            firstname = fullname_parts[-1]
            surname = " ".join(fullname_parts[:-1])
        else:
            surname = fullname_parts[0]
            firstname = ""

        email = ""

        base_filename = f"{class_name}_{surname}_{firstname}.jpg"

        variations = set()
        for sn in [surname, surname.replace(" ", "_"), surname.replace("_", " ")]:
            for fn in [firstname, firstname.replace(" ", "_"), firstname.replace("_", " ")]:
                variations.add(f"{class_name}_{sn}_{fn}.jpg")

        image_path = None
        for cand in variations:
            cand_path = os.path.join(photo_dir, cand)
            if os.path.exists(cand_path):
                image_path = cand_path
                break

        if not image_path:
            try:
                for fname in os.listdir(photo_dir):
                    low = fname.lower()
                    if (class_name.lower() in low
                        and surname.lower().replace(" ", "").replace("_", "") in low.replace(" ", "").replace("_", "")
                        and firstname.lower().replace(" ", "").replace("_", "") in low.replace(" ", "").replace("_", "")):
                        image_path = os.path.join(photo_dir, fname)
                        break
            except FileNotFoundError:
                pass

        if not image_path:
            image_path = os.path.join(photo_dir, base_filename)

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

            # Förnamn
            fontsize_fn = 120
            c.setFont("Helvetica-Bold", fontsize_fn)
            while c.stringWidth(firstname, "Helvetica-Bold", fontsize_fn) > max_width and fontsize_fn > 10:
                fontsize_fn -= 1
                c.setFont("Helvetica-Bold", fontsize_fn)
            c.drawCentredString(page_w/2, page_h * y_positions[0], firstname)

            # Efternamn
            fontsize_sn = 96
            c.setFont("Helvetica-Bold", fontsize_sn)
            while c.stringWidth(surname, "Helvetica-Bold", fontsize_sn) > max_width and fontsize_sn > 10:
                fontsize_sn -= 1
                c.setFont("Helvetica-Bold", fontsize_sn)
            c.drawCentredString(page_w/2, page_h * y_positions[1], surname)

            # Klass
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
    c.setFillColor(colors.black)
    c.drawString(x_start, y_pos, text_before)
    x_cursor = x_start + c.stringWidth(text_before, "Helvetica", 8)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(x_cursor, y_pos, timestamp)
    x_cursor += c.stringWidth(timestamp, "Helvetica-Bold", 8)
    c.setFont("Helvetica", 8)
    c.drawString(x_cursor, y_pos, text_after + author_text + ".")
    author_width = c.stringWidth(author_text, "Helvetica", 8)
    c.linkURL(mailto, (x_cursor + c.stringWidth(text_after, "Helvetica", 8),
                        y_pos - 2,
                        x_cursor + c.stringWidth(text_after, "Helvetica", 8) + author_width,
                        y_pos + 8), relative=0, thickness=0, color=colors.blue)
    c.setFillColor(colors.black)
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
        top_margin, bottom_margin = 30*mm, 30*mm
        left_margin, right_margin = 25*mm, 25*mm

        usable_height = page_h - top_margin - bottom_margin
        usable_width = page_w - left_margin - right_margin

        orig_cell_w = usable_width / cols
        orig_cell_h = usable_height / rows_per_page

        margin_factor = 0.9
        cell_w = orig_cell_w * margin_factor
        cell_h = orig_cell_h * margin_factor

        photo_w = cell_w * 0.9
        photo_h = cell_h * 0.8

        students_per_page = cols * rows_per_page
        total_pages = (len(rows) + students_per_page - 1) // students_per_page

        def draw_header():
            c.setFont("Helvetica-Bold", 36)
            c.drawCentredString(page_w/2, page_h - 20*mm, class_name)
            student_count_text = f"Klassen har {len(rows)} elever."
            c.setFont("Helvetica", 12)
            c.drawCentredString(page_w/2, page_h - 27*mm, student_count_text)

        for page in range(total_pages):
            draw_header()
            page_students = rows[page * students_per_page:(page + 1) * students_per_page]

            for idx in range(students_per_page):
                row_idx = idx // cols
                col_idx = idx % cols

                cell_x = left_margin + col_idx * orig_cell_w + (orig_cell_w - cell_w) / 2
                cell_y = page_h - top_margin - (row_idx + 1) * orig_cell_h + (orig_cell_h - cell_h) / 2

                if idx < len(page_students):
                    r = page_students[idx]
                    firstname, surname, image_path = r["firstname"], r["surname"], r["@image"]

                    # Rita bild
                    has_image = False
                    if image_path and os.path.exists(image_path):
                        try:
                            img_reader = get_optimized_image_reader(image_path, photo_w, photo_h, quality=quality)
                            if img_reader:
                                img = ImageReader(img_reader)
                                img_w, img_h = img.getSize()
                                img_aspect = img_w / img_h
                                frame_aspect = photo_w / photo_h

                                if img_aspect > frame_aspect:
                                    draw_w = photo_w
                                    draw_h = photo_w / img_aspect
                                else:
                                    draw_h = photo_h
                                    draw_w = photo_h * img_aspect

                                draw_x = cell_x + (cell_w - draw_w) / 2
                                draw_y = cell_y + (cell_h - draw_h) / 2

                                c.drawImage(img_reader, draw_x, draw_y, width=draw_w, height=draw_h)
                                has_image = True
                        except Exception:
                            pass

                    # Namn under bilden
                    c.setFont("Helvetica-Bold", 14)
                    name_y = cell_y - 14
                    c.drawCentredString(cell_x + cell_w/2, name_y, f"{firstname} {surname}")

                    # Ram
                    frame_color = colors.black if has_image else colors.red
                    c.setStrokeColor(frame_color)
                    c.rect(cell_x, cell_y, cell_w, cell_h, stroke=1, fill=0)
                    c.setStrokeColor(colors.black)

                    # Text om bild saknas
                    if not has_image:
                        c.setFont("Helvetica-Bold", 14)
                        c.setFillColor(colors.red)
                        c.drawCentredString(cell_x + cell_w/2, cell_y + cell_h/2, "Bild saknas")
                        c.setFillColor(colors.black)

                        msg = f"Saknad bild för {firstname} {surname}: {os.path.normpath(image_path)}"
                        write_log_line(log_file, "SAKNAS", msg)
                        missing += 1
                        missing_list.append(r)

            draw_footer_with_mail(c, page_w, page_h, page + 1, total_pages, version)
            if page < total_pages - 1:
                c.showPage()

        c.save()
        write_log_line(log_file, "INFO", f"PDF Klassfoton skapad: {os.path.normpath(pdf_path)} (saknade bilder: {missing})")
        return missing
    except Exception as e:
        write_log_line(log_file, "ERROR", f"Fel vid skapande av klassfoton för {class_name}: {e}")
        return missing

# ===== 14. HUVUDLOOP =====
total_missing_global = 0
for base_folder in base_folders:
    PHOTO_DIR = os.path.join(base_folder, "Foton")
    CSV_DIR = os.path.join(base_folder, "CSV")
    OUT_BADGES = os.path.join(base_folder, "Namnskyltar")
    OUT_CLASS = os.path.join(base_folder, "Klassfoton")
    TXT_DIR = os.path.join(base_folder, "Elevdata")
    os.makedirs(OUT_CLASS, exist_ok=True)

    datum = datetime.now().strftime("%Y%m%d")
    LOG_FILE = os.path.join(base_folder, f"loggfil_{datum}.txt")
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
