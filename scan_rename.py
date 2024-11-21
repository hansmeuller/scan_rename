import os
import pytesseract
from pdf2image import convert_from_path
import re
from datetime import datetime


scans_folder = os.path.dirname(os.path.realpath(__file__))
log_file_path = os.path.join(scans_folder, "Logeinträge.txt")


PAGE_HEIGHT = 838.0
PAGE_WIDTH = 594.0
PLZ_REGEX = re.compile(r"\b\d{5}\b")
IGNORE_WORDS = {"Postfach", "Postzentrum"}
KEYWORDS = ["Bitte bei Antwort angeben", "Aktenzeichen"]

# Logging-Funktion
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"{timestamp} - {message}\n")

# OCR
def perform_ocr(image):
    try:
        text = pytesseract.image_to_string(image, lang="deu")
        return text.splitlines()
    except Exception as e:
        log_message(f"OCR-Fehler: {e}")
        return []

# extrahieren
def extract_sender(lines):
    for line in lines:
        if PLZ_REGEX.search(line) and len(line.split(",")) > 1:
            absender_text = line.split(",")[0].strip()
            if any(ignore in absender_text for ignore in IGNORE_WORDS):
                continue
            return absender_text[:30]
    return "Kein Absender gefunden"

# Betreff extrahieren
def extract_subject(file_name, lines, page_height):
    # Prüfen
    name_parts = file_name.replace(".pdf", "").split("_")
    for part in name_parts:
        if len(part.split()) > 2:  # zwei Worte
            return " ".join(part.split()[:4])  # Max 4 Worte

    # < 400 "Kontoauszug"
    if page_height < 400:
        return "Kontoauszug"

    # Text extrahieren
    for idx, line in enumerate(lines):
        if 7.5 / PAGE_HEIGHT * 100 <= idx / len(lines) * 100 <= 8.5 / PAGE_HEIGHT * 100:
            return " ".join(line.split()[:5])  # Maximal 5 Worte
    return "Kein Betreff gefunden"

# extrahieren
def extract_case_number(lines):
    for idx, line in enumerate(lines):
        if any(keyword in line for keyword in KEYWORDS):
            if idx + 2 < len(lines):
                case_line = lines[idx + 2].strip()
                if len(case_line) <= 10:
                    return case_line
    return "Kein Aktenzeichen gefunden"

# Extraktion
def process_pdf(file_path):
    try:
        file_name = os.path.basename(file_path)
        images = convert_from_path(file_path, dpi=300, first_page=1, last_page=1)
        if not images:
            log_message(f"Keine Bilder extrahiert für {file_path}")
            return

        text_lines = perform_ocr(images[0])

        # Absender extrahieren
        sender = extract_sender(text_lines)
        log_message(f"Gefundener Absender in {file_path}: {sender}")

        # Betreff extrahieren
        subject = extract_subject(file_name, text_lines, images[0].height)
        log_message(f"Gefundener Betreff in {file_path}: {subject}")

        # Az extrahieren
        case_number = extract_case_number(text_lines)
        log_message(f"Gefundenes Aktenzeichen in {file_path}: {case_number}")

    except Exception as e:
        log_message(f"Fehler bei der Verarbeitung von {file_path}: {e}")

# PDF-Verarbeitung
def process_pdfs(scans_folder):
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Starte Verarbeitung für Datei: {file_path}")
            process_pdf(file_path)

# Main
if __name__ == "__main__":
    process_pdfs(scans_folder)
