import time
import os
import re
import pdfminer
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
from datetime import datetime, timedelta
from pdfminer.high_level import extract_text
from pdfminer.layout import LAParams

# scant current
scans_folder = os.path.dirname(os.path.realpath(__file__))
# current Verzeichnis
log_file_path = os.path.join(scans_folder, "Logeinträge_ScanRename.txt")
log_file_exists = False
print(f"Logdatei-pfad: {log_file_path}")


# Logdatei prüfen/create
def check_and_create_log_file():
    """create if not exist"""
    global log_file_exists
    print(f"prüfe ob logdatei existiert pfad: {log_file_path}")
    if not os.path.exists(log_file_path):
        print("erstelle logdatei")
        with open(log_file_path, "w") as log_file:
            print("log erstellt")
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    else:
        print("exist")
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


# extrahiere inklusive OCR-Sicherung
def extract_text_with_format_from_pdf(pdf_path):
    """Extrahiert"""
    formatted_text = []

    try:
        text = extract_text(pdf_path, laparams=LAParams())
        if text:
            lines = text.split("\n")
            for line in lines:
                formatted_text.append((line.strip(), 'normal'))
    except Exception as e:
        log_message(f"Fehler beim Öffnen der Datei {pdf_path}: {e}")

    # OCR
    if not formatted_text:
        log_message(f"Keine Textinhalte gefunden, führe OCR für Datei {pdf_path} aus...")
        try:
            images = convert_from_path(pdf_path)
            for image in images:
                text = pytesseract.image_to_string(image, lang='deu')
                if text.strip():
                    formatted_text.append((text.strip(), 'normal'))
        except Exception as e:
            log_message(f"Fehler bei der OCR-Verarbeitung von {pdf_path}: {e}")

    return formatted_text


# suche fettgedrucktes
def get_subject_from_formatted_text(formatted_text):
    """fettgedrucktes"""
    subject = None

    # Zuerst
    for line, format in formatted_text:
        if format == "bold":
            subject = line.strip()
            break

    # az
    if not subject:
        for line, format in formatted_text:
            if re.search(r'aktenzeichen[:\s]*[a-zA-Z0-9\/\-\.]+', line, re.IGNORECASE):
                subject = f"Aktenzeichen_{line.strip()}"
                break

    return subject


# Datei umbenennen
def rename_file(old_path, absender, date, subject):
    """Absender_Datum_Betreff"""
    directory, old_file_name = os.path.split(old_path)
    file_extension = old_file_name.split('.')[-1]
    new_file_name = sanitize_filename(f"{absender}_{date}_{subject}.{file_extension}")
    new_file_path = os.path.join(directory, new_file_name)

    if not os.path.exists(new_file_path):
        os.rename(old_path, new_file_path)
        log_message(f"Datei umbenannt von {old_file_name} zu {new_file_name}")
    else:
        log_message(f"Fehler: Die Datei {new_file_path} existiert bereits.")


# bereinigen
def sanitize_filename(filename):
    """Entfernt"""
    return re.sub(r'[\/\:*?"<>|]', '_', filename)


# Hauptprozess
def process_scan_files(scans_folder):
    """bearbeitet"""
    log_message(f"Arbeitsverzeichnis: {os.getcwd()}")
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            # Extrahiere
            formatted_text = extract_text_with_format_from_pdf(file_path)
            subject = get_subject_from_formatted_text(formatted_text)

            if subject:
                absender = file_name.split('_')[1]  # zweiter Stelle
                date = file_name.split('_')[0]  # Datum an erster Stelle
                rename_file(file_path, absender, date, subject)
            else:
                log_message(f"Kein Betreff gefunden für {file_name}.")


# Hauptverarbeitung
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