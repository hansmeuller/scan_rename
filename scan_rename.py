import os
import re
import fitz  # PyMuPDF
import ocrmypdf
import time
from datetime import datetime

# current
scans_folder = os.path.dirname(os.path.realpath(__file__))
log_file_path = os.path.join(scans_folder, "Logeinträge_ScanRename.txt")

# headline
KNICKFALTE_Y_RATIO = 1 / 3


# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# OCR
def ocr_and_save_temp_pdf(pdf_path):
    temp_pdf_path = pdf_path.replace('.pdf', '_ocr_temp.pdf')
    ocrmypdf.ocr(pdf_path, temp_pdf_path, force_ocr=True, output_type="pdfa")
    return temp_pdf_path


# Betreff
def extract_subject_from_pdf(pdf_path):
    try:
        with fitz.open(pdf_path) as pdf_document:
            page = pdf_document.load_page(0)
            page_height = page.rect.height

            # x==ktoauszug?
            if page_height < 500:
                log_message(f"{pdf_path}: Kontoauszug erkannt aufgrund der Seitenhöhe.")
                return "Kontoauszug"

            # Extrahiere
            knickfalte_y = page_height * KNICKFALTE_Y_RATIO
            words = page.get_text("words")

            # Suche
            full_text = page.get_text("text")
            if "Bitte bei Antwort angeben" in full_text:
                lines = full_text.splitlines()
                index = next((i for i, line in enumerate(lines) if "Bitte bei Antwort angeben" in line), None)
                if index is not None and index + 2 < len(lines):
                    aktenzeichen = lines[index + 2][:16]
                    log_message(f"Aktenzeichen erkannt: {aktenzeichen}")
                    return f"Aktenzeichen {aktenzeichen}"

            # first five words
            knick_text = []
            for w in words:
                if w[1] >= knickfalte_y:
                    knick_text.append(w[4])
                    if len(knick_text) >= 5:
                        break
            subject = ' '.join(knick_text)
            log_message(f"Knickfalte als Betreff extrahiert: {subject}")
            return subject

    except Exception as e:
        log_message(f"Fehler beim Extrahieren aus {pdf_path}: {e}")
        return "Betreff unbekannt"


# Datei umbenennen
def rename_file(old_path, date, subject):
    directory, old_file_name = os.path.split(old_path)
    old_file_name = old_file_name.replace("\u2069", "")  # Unicode-Steuerzeichen entfernen
    file_extension = old_file_name.split('.')[-1]
    new_file_name = f"Unbekannt_{date}_{subject}.{file_extension}"
    new_file_path = os.path.join(directory, new_file_name)

    if old_path != new_file_path:
        if not os.path.exists(new_file_path):
            os.rename(old_path, new_file_path)
            log_message(f"Datei umbenannt von {old_file_name} zu {new_file_name}")
        else:
            log_message(f"Fehler: Die Datei {new_file_path} existiert bereits.")


# mainmain
def process_pdfs(scans_folder):
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf") and not file_name.startswith("._"):
            file_path = os.path.join(scans_folder,
                                     file_name.replace("\u2069", ""))  # Steuerzeichen aus Dateipfad entfernen
            log_message(f"Verarbeite {file_name}...")

            # current date
            date_match = re.match(r'\d{8}', file_name)
            date = date_match.group(0) if date_match else datetime.now().strftime('%Y%m%d')

            # OCR
            temp_pdf_path = ocr_and_save_temp_pdf(file_path)

            # extrahieren
            subject = extract_subject_from_pdf(temp_pdf_path)

            # umbenennen
            rename_file(file_path, date, subject)

            # Temporär
            os.remove(temp_pdf_path)


# main
def main():
    try:
        log_message("Start erfolgreich")
        process_pdfs(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")
        time.sleep(5)
        main()


if __name__ == "__main__":
    main()
