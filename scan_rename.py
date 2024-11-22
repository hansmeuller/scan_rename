import os
import numpy as np
from pdf2image import convert_from_path
import easyocr
import re
from datetime import datetime

# initialisieren
reader = easyocr.Reader(["de"], gpu=False)

# current
scans_folder = os.path.dirname(os.path.realpath(__file__))
log_file_path = os.path.join(scans_folder, "Logeinträge.txt")

# Parameter
PAGE_HEIGHT = 838.0
PAGE_WIDTH = 594.0
PLZ_REGEX = re.compile(r"\b\d{5}\b")
IGNORE_WORDS = {"Postfach", "Postzentrum"}
KEYWORDS = ["Bitte bei Antwort angeben", "Aktenzeichen"]


# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# extraktion
def perform_ocr(image):
    try:

        image = image.convert("RGB")
        image_np = np.array(image)

        results = reader.readtext(image_np, detail=0, paragraph=True)
        return results

    except Exception as e:
        log_message(f"OCR-Fehler: {e}")
        return []


# Absender extrahieren
def extract_sender(lines):
    for line in lines:
        if PLZ_REGEX.search(line):
            absender_text = re.split(r"[;,]", line)[0].strip()
            if any(ignore in absender_text for ignore in IGNORE_WORDS):
                continue
            log_message(f"Absender-Zeile gefunden: {line.strip()}")
            return absender_text[:30]
    log_message(f"Keine passende Absender-Zeile gefunden in: {lines}")
    return "Kein Absender gefunden"


# Betreff extrahieren
def extract_subject(file_name, lines, page_height):
    name_parts = file_name.replace(".pdf", "").split("_")
    for part in name_parts:
        if len(part.split()) > 2:
            log_message(f"Betreff aus Dateiname: {part}")
            return " ".join(part.split()[:4])

    if page_height < 400:
        log_message(f"Seitenhöhe < 400 erkannt, Betreff: Kontoauszug")
        return "Kontoauszug"

    for idx, line in enumerate(lines):
        if 7.5 / PAGE_HEIGHT * 100 <= idx / len(lines) * 100 <= 8.5 / PAGE_HEIGHT * 100:
            log_message(f"Betreff-Zeile gefunden: {line.strip()}")
            return " ".join(line.split()[:5])
    log_message(f"Keine passende Betreff-Zeile gefunden in: {lines}")
    return "Kein Betreff gefunden"


# az
def extract_case_number(lines):
    for idx, line in enumerate(lines):
        if any(keyword in line for keyword in KEYWORDS):
            if idx + 2 < len(lines):
                case_line = lines[idx + 2].strip()
                if len(case_line) <= 10:
                    log_message(f"Aktenzeichen-Zeile gefunden: {case_line}")
                    return case_line
    log_message(f"Keine passende Aktenzeichen-Zeile gefunden in: {lines}")
    return "Kein Aktenzeichen gefunden"


# Verarbeitung
def process_pdf(file_path):
    try:
        file_name = os.path.basename(file_path)
        images = convert_from_path(file_path, dpi=300, first_page=1, last_page=1)
        if not images:
            log_message(f"Keine Bilder extrahiert für {file_path}")
            return

        text_lines = perform_ocr(images[0])

        sender = extract_sender(text_lines)
        log_message(f"Gefundener Absender in {file_path}: {sender}")

        subject = extract_subject(file_name, text_lines, images[0].height)
        log_message(f"Gefundener Betreff in {file_path}: {subject}")

        case_number = extract_case_number(text_lines)
        log_message(f"Gefundenes Aktenzeichen in {file_path}: {case_number}")

    except Exception as e:
        log_message(f"Fehler bei der Verarbeitung von {file_path}: {e}")


# current verarbeiten
def process_pdfs(scans_folder):
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Starte Verarbeitung für Datei: {file_path}")
            process_pdf(file_path)


# Main
if __name__ == "__main__":
    process_pdfs(scans_folder)
