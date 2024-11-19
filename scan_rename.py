import os
import fitz  # PyMuPDF
import ocrmypdf
import shutil
from datetime import datetime
import re

# current
scans_folder = os.path.dirname(os.path.realpath(__file__))
log_file_path = os.path.join(scans_folder, "Logeinträge.txt")


PAGE_HEIGHT = 838.0
PAGE_WIDTH = 594.0

# Positionen
ABSENDER_TOP_START = 3 / 2.54 * 96  # 3 cm von oben
ABSENDER_TOP_END = 4 / 2.54 * 96  # 4 cm von oben
BETREFF_TOP = 7.5 / 2.54 * 96  # 7,5 cm von oben
KNICKFALTE_TOP = 6.5 / 2.54 * 96  # 6,5 cm von oben
AKTENZEICHEN_LEFT = 6 / 2.54 * 96  # 6 cm von links

# nope
IGNORE_WORDS = {"Postanschrift", "Postzentrum"}
IGNORE_REGEX = re.compile(r"^\d{5}$")  # PLZ (5 Ziffern)

# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"{timestamp} - {message}\n")

# OCR
def apply_ocr_to_pdf(file_path):
    try:
        ocrmypdf.ocr(file_path, file_path, force_ocr=True)
        log_message(f"OCR erfolgreich angewendet auf {file_path}")
    except Exception as e:
        log_message(f"Fehler beim Anwenden von OCR auf {file_path}: {e}")

# Entschlüsseln
def decrypt_pdf(file_path):
    decrypted_path = file_path.replace(".pdf", "_temp_decrypted.pdf")
    try:
        command = ["qpdf", "--decrypt", file_path, decrypted_path]
        result = os.system(" ".join(command))
        if result == 0:
            shutil.move(decrypted_path, file_path)
            log_message(f"PDF erfolgreich entschlüsselt: {file_path}")
        else:
            raise Exception("Entschlüsselung fehlgeschlagen.")
    except Exception as e:
        log_message(f"Fehler beim Entschlüsseln von {file_path}: {e}")
        if os.path.exists(decrypted_path):
            os.remove(decrypted_path)

# az extrahieren
def extract_aktenzeichen(words):
    for i, word in enumerate(words):
        if word[4].lower() == "bitte" and i + 2 < len(words):
            if words[i + 1][4].lower() == "bei" and words[i + 2][4].lower() == "antwort":
                line_y = words[i + 2][3]
                aktenzeichen = []
                for next_word in words[i + 3 :]:
                    if next_word[1] > line_y + 10:  # 2 Zeilen darunter
                        aktenzeichen.append(next_word[4])
                        if len("".join(aktenzeichen).replace(" ", "")) >= 10:
                            return "".join(aktenzeichen)[:10]
    return None

# abs spezial
def extract_sender(words):
    absender = []
    found_plz_line = None
    for w in words:
        if ABSENDER_TOP_START <= w[1] <= ABSENDER_TOP_END and not IGNORE_REGEX.match(w[4]) and w[4] not in IGNORE_WORDS:
            absender.append(w[4])


        if w[1] <= KNICKFALTE_TOP / 2 and IGNORE_REGEX.match(w[4]):
            if found_plz_line is None or w[1] < found_plz_line:
                found_plz_line = w[1]
                absender = [w[4]]

    absender_text = " ".join(absender) if absender else "Keine Absenderangabe erkennbar"
    log_message(f"Gefundene mögliche Absender: {absender_text}")
    return absender_text

# betreff extrahieren
def extract_subject(words):
    betreff = []
    for w in words:
        if abs(w[1] - BETREFF_TOP) <= 10 and w[0] < PAGE_WIDTH / 2:  # Linksseitig begrenzt
            betreff.append(w[4])
    return " ".join(betreff) if betreff else "Kein Betreff gefunden"

# extrahieren
def extract_sender_and_subject(file_path):
    try:
        with fitz.open(file_path) as pdf_document:
            page = pdf_document.load_page(0)
            words = page.get_text("words")

            # Absender
            absender_text = extract_sender(words)

            # extrahieren
            if page.rect.height < 400:  # Kontoauszug
                betreff_text = "Kontoauszug"
            else:
                betreff_text = extract_subject(words)

            log_message(f"Gefundener Betreff in {file_path}: {betreff_text}")

            # Az
            aktenzeichen = extract_aktenzeichen(words)
            if aktenzeichen:
                log_message(f"Gefundenes Aktenzeichen in {file_path}: {aktenzeichen}")

    except Exception as e:
        log_message(f"Fehler beim Extrahieren von Absender und Betreff aus {file_path}: {e}")

# umbenennen
def rename_file(file_path, absender, betreff, datum):
    try:
        absender_text = "_".join(absender.split()) if absender else "Unbekannt"
        betreff_text = "_".join(betreff.split()) if betreff else "Unbekannt"
        new_file_name = f"{absender_text}_{datum}_{betreff_text}.pdf"
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        os.rename(file_path, new_file_path)
        log_message(f"Datei umbenannt in: {new_file_path}")
    except Exception as e:
        log_message(f"Fehler beim Umbenennen von {file_path}: {e}")

# mainmain
def process_pdfs(scans_folder):
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Starte Verarbeitung für Datei: {file_path}")

            # Entschlüsseln
            decrypt_pdf(file_path)

            # OCR
            apply_ocr_to_pdf(file_path)

            # Absender und Betreff
            extract_sender_and_subject(file_path)

# main
if __name__ == "__main__":
    process_pdfs(scans_folder)
