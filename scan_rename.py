import os
import fitz  # PyMuPDF
import ocrmypdf
from datetime import datetime
import re

# Pfad
scans_folder = os.path.dirname(os.path.realpath(__file__))
# Pfad
log_file_path = os.path.join(scans_folder, "Logeinträge_TestExtraction.txt")

# Maße der DIN A4-Seiten in Pixel
PAGE_HEIGHT = 838.0
PAGE_WIDTH = 594.0

# Positionen in Pixel
ABSENDER_TOP_START = (3 / 2.54) * 96  # 3 cm von oben
ABSENDER_TOP_END = (3.5 / 2.54) * 96  # 3,5 cm von oben
BETREFF_TOP = (7.5 / 2.54) * 96  # 7,5 cm von oben
KNICKFALTE_Y_POSITION = (6.5 / 2.54) * 96  # 6,5 cm von oben
AKTENZEICHEN_LEFT = (6 / 2.54) * 96  # 6 cm von links, mittig


# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# OCR
def apply_ocr_to_pdf(pdf_path):
    try:
        ocrmypdf.ocr(pdf_path, pdf_path, force_ocr=True)
        log_message(f"OCR erfolgreich angewendet auf {pdf_path}")
    except Exception as e:
        log_message(f"Fehler beim Anwenden von OCR auf {pdf_path}: {e}")


# Absender, Betreff
def extract_sender_and_subject(pdf_path):
    try:
        with fitz.open(pdf_path) as pdf_document:
            page = pdf_document.load_page(0)  # Nur erste Seite
            page_height = page.rect.height

            # Seitenhöhe
            if page_height < 400:
                log_message(f"Kontoauszug erkannt in {pdf_path}")
                absender = extract_sender(page)
                if absender:
                    log_message(f"Gefundener Absender in {pdf_path}: {' '.join(absender)}")
                return

            words = page.get_text("words")
            absender = extract_sender(page)
            betreff = extract_betreff(words)
            aktenzeichen = extract_aktenzeichen(words)

            if absender:
                log_message(f"Gefundener Absender in {pdf_path}: {' '.join(absender)}")
            if betreff:
                log_message(f"Gefundener Betreff in {pdf_path}: {' '.join(betreff)}")
            if aktenzeichen:
                log_message(f"Gefundenes Aktenzeichen in {pdf_path}: {aktenzeichen}")

    except Exception as e:
        log_message(f"Fehler beim Extrahieren des Absenders, Betreffs und Aktenzeichens aus {pdf_path}: {e}")


# Absender extrahieren
def extract_sender(page):
    words = page.get_text("words")
    absender = []
    ignore_list = ["Postanschrift", "Postzentrum"]
    for w in words:
        if ABSENDER_TOP_START <= w[1] <= ABSENDER_TOP_END:
            if not re.match(r'\d{5}', w[4]) and w[4] not in ignore_list:  # Wörter ignorieren
                absender.append(w[4])
    return absender


# Betreff
def extract_betreff(words):
    betreff = []
    for w in words:
        if abs(w[1] - BETREFF_TOP) <= 5:
            betreff.append(w[4])
    return betreff[:5]  # 5 Wörter


# Aktenzeichen extrahieren
def extract_aktenzeichen(words):
    aktenzeichen = []
    found_aktenzeichen = False
    for i, w in enumerate(words):
        if w[4].lower() == "bitte" and i + 3 < len(words) and words[i + 1][4].lower() == "bei" and words[i + 2][
            4].lower() == "antwort" and words[i + 3][4].lower() == "angeben":
            # nächste Zeile nach dem Schlagwort
            line_y = words[i + 3][3]
            for next_word in words[i + 4:]:
                if next_word[1] > line_y + 10:
                    aktenzeichen.append(next_word[4])
                    if len(''.join(aktenzeichen).replace(" ", "")) >= 8:
                        found_aktenzeichen = True
                        break
        if found_aktenzeichen:
            break
    aktenzeichen_text = ''.join(aktenzeichen)[:10]  # 10 Zeichen
    return aktenzeichen_text


# main Verarbeitung
def process_pdfs(scans_folder):
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf") and not file_name.startswith("._"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Starte Verarbeitung für Datei: {file_path}")
            apply_ocr_to_pdf(file_path)  # OCR auf die PDF anwenden
            extract_sender_and_subject(file_path)
            log_message(f"Verarbeitung abgeschlossen für Datei: {file_path}")


# main
def main():
    try:
        log_message("Start erfolgreich")
        process_pdfs(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")


if __name__ == "__main__":
    main()
