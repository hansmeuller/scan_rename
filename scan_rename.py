import os
import fitz  # PyMuPDF
import ocrmypdf
from datetime import datetime, timedelta
import re

# current path
scans_folder = os.path.dirname(os.path.realpath(__file__))
# current log
log_file_path = os.path.join(scans_folder, "Logeinträge_TestExtraction.txt")

# Maße DIN A4-Seiten
PAGE_HEIGHT = 838.0
PAGE_WIDTH = 594.0

# Positionen
ABSENDER_TOP_START = (3 / 2.54) * 96  # 3 cm von oben
ABSENDER_TOP_END = (4 / 2.54) * 96  # 4 cm von oben
BETREFF_TOP = (7.5 / 2.54) * 96  # 7,5 cm von oben
KNICKFALTE_Y_POSITION = (6.5 / 2.54) * 96  # 6,5 cm von oben
AKTENZEICHEN_LEFT = (6 / 2.54) * 96  # 6 cm von links, mittig


# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# delete <=2tage
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


# OCR
def apply_ocr_to_pdf(pdf_path):
    try:
        ocrmypdf.ocr(pdf_path, pdf_path, force_ocr=True)
        log_message(f"OCR erfolgreich angewendet auf {pdf_path}")
    except Exception as e:
        log_message(f"Fehler beim Anwenden von OCR auf {pdf_path}: {e}")


# absender betreff extrahieren
def extract_sender_and_subject(pdf_path):
    try:
        with fitz.open(pdf_path) as pdf_document:
            page = pdf_document.load_page(0)  # 1 seite only
            page_height = page.rect.height

            # Kontoauszugserkennung
            if page_height < 400:
                log_message(f"Kontoauszug erkannt in {pdf_path}")
                absender = extract_sender(page)
                if absender:
                    log_message(f"Gefundener Absender in {pdf_path}: {' '.join(absender)}")
                new_file_name = rename_file(file_path=pdf_path, absender=absender, betreff=["Kontoauszug"],
                                            datum=extract_date_from_filename(pdf_path))
                log_message(f"Datei umbenannt in: {new_file_name}")
                return

            words = page.get_text("words")
            absender = extract_sender(page)
            betreff = []
            if len(re.findall(r'\w+', os.path.basename(pdf_path))) > 2:
                betreff = re.findall(r'\w+', os.path.basename(pdf_path))[:4]  # 4 Wörter
            else:
                betreff = extract_betreff(words)
            aktenzeichen = extract_aktenzeichen(words)

            if absender:
                log_message(f"Gefundener Absender in {pdf_path}: {' '.join(absender)}")
            if betreff:
                log_message(f"Gefundener Betreff in {pdf_path}: {' '.join(betreff)}")
            if aktenzeichen:
                log_message(f"Gefundenes Aktenzeichen in {pdf_path}: {aktenzeichen}")
                betreff = [aktenzeichen]  # if az==betreff

            # Datei umbenennen
            new_file_name = rename_file(file_path=pdf_path, absender=absender, betreff=betreff,
                                        datum=extract_date_from_filename(pdf_path))
            log_message(f"Datei umbenannt in: {new_file_name}")

    except Exception as e:
        log_message(f"Fehler beim Extrahieren des Absenders, Betreffs und Aktenzeichens aus {pdf_path}: {e}")


# Absender extrahieren
def extract_sender(page):
    words = page.get_text("words")
    absender = []
    ignore_list = ["Postanschrift", "Postzentrum"]
    for w in words:
        if ABSENDER_TOP_START <= w[1] <= ABSENDER_TOP_END:
            if not re.match(r'\d{5}', w[4]) and w[4] not in ignore_list:  # PLZ und bestimmte Wörter ignorieren
                absender.append(w[4])
    return absender


# Betreff extrahieren
def extract_betreff(words):
    betreff = []
    for w in words:
        if abs(w[1] - BETREFF_TOP) <= 10:
            betreff.append(w[4])
    return betreff[:5]  # 5 Wörter


# extrahieren 8 Zeichen begrenzt
def extract_aktenzeichen(words):
    aktenzeichen = []
    found_aktenzeichen = False
    for i, w in enumerate(words):
        if w[4].lower() == "bitte" and i + 3 < len(words) and words[i + 1][4].lower() == "bei" and words[i + 2][
            4].lower() == "antwort" and words[i + 3][4].lower() == "angeben":
            # Extrahiere
            line_y = words[i + 3][3]  # Y-Position
            for next_word in words[i + 4:]:
                if next_word[1] > line_y + 10:  # Nächste Zeile
                    aktenzeichen.append(next_word[4])
                    if len(''.join(aktenzeichen).replace(" ", "")) >= 8:
                        found_aktenzeichen = True
                        break
        if found_aktenzeichen:
            break
    aktenzeichen_text = ''.join(aktenzeichen)[:10]  # 10 Zeichen
    return aktenzeichen_text


# Extrahieren
def extract_date_from_filename(file_name):
    date_match = re.search(r'\d{8}', file_name)
    if date_match:
        return date_match.group(0)
    return ""


# Datei umbenennen
def rename_file(file_path, absender, betreff, datum):
    if isinstance(betreff, list):
        betreff = betreff[:4]  # 4 Wörter
    absender_text = '_'.join(absender) if absender else "UnknownSender"
    betreff_text = '_'.join(betreff) if betreff else "UnknownSubject"
    new_file_name = f"{absender_text}_{datum}_{betreff_text}.pdf"
    new_file_name = re.sub(r'[^a-zA-Z0-9_]', '', new_file_name)  # Entfernt
    new_file_name = re.sub(r'_+', '_', new_file_name)
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
    if not os.path.exists(new_file_path) and file_path != new_file_path:  # umbenennen if not exist
        os.rename(file_path, new_file_path)
    return new_file_name


# mainmain
def process_pdfs(scans_folder):
    if not os.path.exists(log_file_path):
        with open(log_file_path, "a") as log_file:
            log_file.write("Logdatei erstellt\n")
    clean_old_logs()
    processed_files = set()  # verarbeitet
    for root, _, files in os.walk(scans_folder):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                if pdf_path in processed_files:
                    continue
                log_message(f"Starte Verarbeitung für Datei: {pdf_path}")
                apply_ocr_to_pdf(pdf_path)
                extract_sender_and_subject(pdf_path)
                processed_files.add(pdf_path)

            # main


if __name__ == "__main__":
    process_pdfs(scans_folder)
