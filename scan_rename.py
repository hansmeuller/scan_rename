import time
import os
import re
import ocrmypdf
import fitz  # PyMuPDF
from datetime import datetime, timedelta

# scant current
scans_folder = os.path.dirname(os.path.realpath(__file__))
# Pfad current
log_file_path = os.path.join(scans_folder, "Logeinträge_ScanRename.txt")
log_file_exists = False


# Logdatei prüfen/create
def check_and_create_log_file():
    """create if not exist"""
    global log_file_exists
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    log_file_exists = True


# Einträge mit timestamp
def log_message(message):
    """message+timestamp"""
    check_and_create_log_file()
    clean_old_log_entries()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# lösche +=woche
def clean_old_log_entries():
    """Löscht einträge +=woche"""
    if os.path.exists(log_file_path):
        with open(log_file_path, "r") as log_file:
            lines = log_file.readlines()

        woche_alt = datetime.now() - timedelta(days=7)
        with open(log_file_path, "w") as log_file:
            for line in lines:
                try:
                    log_time = datetime.strptime(line.split(" - ")[0], "%Y-%m-%d %H:%M:%S")
                    if log_time > woche_alt:
                        log_file.write(line)
                except (ValueError, IndexError):
                    log_file.write(line)


# OCR
def ocr_and_extract_text(pdf_path, retry_count=0):
    """Extrahiert"""
    text = ""

    try:
        # OCR
        temp_txt_file = pdf_path.replace('.pdf', '_text.txt')
        ocrmypdf.ocr(pdf_path, pdf_path, force_ocr=True, sidecar=temp_txt_file)
        with open(temp_txt_file, 'r') as txt_file:
            text = txt_file.read()
        os.remove(temp_txt_file)

        log_message(f"Text erfolgreich extrahiert für Datei {pdf_path}")
        log_message(f"Extrahierter Text: {text[:500]}...")  # 500 Zeichen
    except Exception as e:
        if retry_count < 3:
            log_message(f"Fehler bei der Textverarbeitung von {pdf_path}: {e}. Wiederhole ({retry_count + 1}/3)")
            time.sleep(5)
            return ocr_and_extract_text(pdf_path, retry_count=retry_count + 1)
        else:
            log_message(f"Fehler bei der Textverarbeitung von {pdf_path} nach mehreren Versuchen: {e}")
            return ""

    return text


# Suche
def extract_bold_text_from_pdf(pdf_path):
    """Extrahiert"""
    bold_text = []
    try:
        with fitz.open(pdf_path) as pdf_document:
            for page_number in range(len(pdf_document)):
                page = pdf_document.load_page(page_number)
                blocks = page.get_text("dict")['blocks']
                for block in blocks:
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            if span.get("flags", 0) & 2:  # Flag 2
                                bold_text.append(span.get("text", "").strip())
    except Exception as e:
        log_message(f"Fehler beim Extrahieren von fettgedrucktem Text aus {pdf_path}: {e}")

    return " ".join(bold_text)


# suche
def get_subject_from_text(text, pdf_path):
    """Sucht"""
    subject = None

    try:
        # Fettgedrucktes
        bold_text = extract_bold_text_from_pdf(pdf_path)
        if bold_text:
            log_message(f"Gefundener fettgedruckter Text: {bold_text}")
            subject = bold_text

        # fettgedruckter Text
        if not subject:
            lines = text.split("\n")
            for line in lines:
                if re.search(r'gericht|beschluss|urteil', line, re.IGNORECASE):
                    subject = line.strip()
                    log_message(f"Gefundener Betreff: {subject}")
                    break
    except Exception as e:
        log_message(f"Fehler beim Verarbeiten des Texts: {e}")

    return subject


# Datei umbenennen
def rename_file(old_path, absender, date, subject):
    """Absender_Datum_Betreff"""
    directory, old_file_name = os.path.split(old_path)
    file_extension = old_file_name.split('.')[-1]
    new_file_name = sanitize_filename(f"{absender}_{date}_{subject}.{file_extension}")
    new_file_path = os.path.join(directory, new_file_name)

    if old_path != new_file_path:
        if not os.path.exists(new_file_path):
            os.rename(old_path, new_file_path)
            log_message(f"Datei umbenannt von {old_file_name} zu {new_file_name}")
        else:
            log_message(f"Fehler: Die Datei {new_file_path} existiert bereits.")


# Dateinamen bereinigen
def sanitize_filename(filename):
    """Entfernt"""
    return re.sub(r'[/\:*?"<>|]', '_', filename)


# Hauptprozess
def process_scan_files(scans_folder):
    """Durchsucht"""
    log_message(f"Arbeitsverzeichnis: {os.getcwd()}")
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf") and not file_name.startswith("._"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            # OCR
            text = ocr_and_extract_text(file_path)
            if not text:
                continue

            subject = get_subject_from_text(text, file_path)

            if subject:
                absender = file_name.split('_')[1]  # Absender an zweiter Stelle im Dateinamen
                date = file_name.split('_')[0]  # Datum an erster Stelle im Dateinamen
                # DatumFormat
                if re.match(r'\d{8}', date):
                    rename_file(file_path, absender, date, subject)
                else:
                    log_message(f"Fehler: Datum im falschen Format für {file_name}")
            else:
                log_message(f"Kein Betreff gefunden für {file_name}.")

            # Loops verhindern
            if "._" in file_name:
                log_message(f"Überspringe Datei {file_name} aufgrund eines möglichen Fehlers im PDF-Format.")
                continue


# Main
def main():
    """Main"""
    try:
        log_message("Start erfolgreich")
        process_scan_files(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")
        time.sleep(5)
        main()


if __name__ == "__main__":
    main()
