import os
import fitz  # PyMuPDF
import ocrmypdf
from datetime import datetime, timedelta
import re
import subprocess
import shutil

# current
scans_folder = os.path.dirname(os.path.realpath(__file__))
# current
log_file_path = os.path.join(scans_folder, "Logeinträge_TestExtraction.txt")


ABSENDER_TOP_START_CM = 3
ABSENDER_TOP_END_CM = 3.5
BETREFF_TOP_CM = 7.5
KNICKFALTE_TOP_CM = 6.5
AKTENZEICHEN_LEFT_CM = 6


def cm_to_points(cm):
    return (cm / 2.54) * 96

# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")

# delete>=2tage
def clean_old_logs():
    if os.path.exists(log_file_path):
        with open(log_file_path, "r") as log_file:
            lines = log_file.readlines()
        cutoff_date = datetime.now() - timedelta(days=2)
        with open(log_file_path, "w") as log_file:
            for line in lines:
                try:
                    log_date = datetime.strptime(line.split(" - ")[0], "%Y-%m-%d %H:%M:%S")
                    if log_date > cutoff_date:
                        log_file.write(line)
                except ValueError:
                    log_file.write(line)

# Verschlüsselung entfernen
def create_decrypted_copy(pdf_path):
    decrypted_path = pdf_path.replace(".pdf", "_decrypted.pdf")
    if not os.path.exists(decrypted_path):
        subprocess.run(["qpdf", "--decrypt", pdf_path, decrypted_path])
    return decrypted_path

# OCR
def apply_ocr_to_pdf(pdf_path):
    try:
        ocrmypdf.ocr(pdf_path, pdf_path, skip_text=True, output_type="pdf")
        log_message(f"OCR erfolgreich angewendet auf {pdf_path}")
    except Exception as e:
        log_message(f"OCR-Fehler für {pdf_path}: {e}")

# Extraktion
def extract_sender_and_subject(pdf_path):
    try:
        with fitz.open(pdf_path) as pdf_document:
            page = pdf_document.load_page(0)  # Nur erste Seite
            words = page.get_text("words")

            # Ktoerkennung
            page_height = page.rect.height
            if page_height < 400:
                log_message(f"Kontoauszug erkannt in {pdf_path}")
                absender = extract_sender(words, page_height)
                log_message(f"Extrahierter Absender in {pdf_path}: {' '.join(absender) if absender else 'Kein Absender gefunden'}")
                log_message(f"Betreff für {pdf_path}: Kontoauszug")
                return

            absender = extract_sender(words, page_height)
            betreff = []
            aktenzeichen = extract_aktenzeichen(words, page_height)

            # Betreff
            if aktenzeichen:
                betreff = [aktenzeichen]
                log_message(f"Gefundenes Aktenzeichen in {pdf_path}: {aktenzeichen}")
            else:
                betreff = extract_betreff(words, page_height)
                log_message(f"Gefundener Betreff in {pdf_path}: {' '.join(betreff) if betreff else 'Kein Betreff gefunden'}")

            log_message(f"Gefundener Absender in {pdf_path}: {' '.join(absender) if absender else 'Kein Absender gefunden'}")

    except Exception as e:
        log_message(f"Fehler beim Extrahieren des Absenders, Betreffs und Aktenzeichens aus {pdf_path}: {e}")

# Extraktion
def extract_sender(words, page_height):
    absender = []
    top_start = cm_to_points(ABSENDER_TOP_START_CM)
    top_end = cm_to_points(ABSENDER_TOP_END_CM)
    ignore_list = ["Postanschrift", "Postzentrum"]
    for w in words:
        if top_start <= w[1] <= top_end and w[4] not in ignore_list:
            absender.append(w[4])
    return absender

# dynamische Position
def extract_betreff(words, page_height):
    betreff = []
    betreff_top = cm_to_points(BETREFF_TOP_CM)
    for w in words:
        if abs(w[1] - betreff_top) <= 10:
            betreff.append(w[4])
    return betreff[:5]  # Maximal 5 Wörter

# Extraktion
def extract_aktenzeichen(words, page_height):
    aktenzeichen = []
    for i, w in enumerate(words):
        if w[4].lower() == "bitte" and i + 3 < len(words) and words[i + 1][4].lower() == "bei" and words[i + 2][4].lower() == "antwort" and words[i + 3][4].lower() == "angeben":
            line_y = words[i + 3][3]
            for next_word in words[i+4:]:
                if next_word[1] > line_y + 10:
                    aktenzeichen.append(next_word[4])
                    if len(''.join(aktenzeichen).replace(" ", "")) >= 10:
                        return ''.join(aktenzeichen)[:10]  # Maximal 10 Zeichen
    return None

# mainmain
def process_pdfs(scans_folder):
    if not os.path.exists(log_file_path):
        with open(log_file_path, "a") as log_file:
            log_file.write("Logdatei erstellt\n")
    clean_old_logs()
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf") and not file_name.startswith("._"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Starte Verarbeitung für Datei: {file_path}")
            decrypted_path = create_decrypted_copy(file_path)
            apply_ocr_to_pdf(decrypted_path)
            extract_sender_and_subject(decrypted_path)

# maiin
if __name__ == "__main__":
    process_pdfs(scans_folder)
