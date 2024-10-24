import time
import os
import re
import ocrmypdf
from datetime import datetime, timedelta

# scant im aktuellen Verzeichnis, in dem das Programm liegt
scans_folder = os.path.dirname(os.path.realpath(__file__))
# Pfad Logdatei im aktuellen Verzeichnis
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
    """+=woche"""
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


# OCR-Verarbeitung und Text Extraktion mit OCRmyPDF
def ocr_and_extract_text(pdf_path, retry_count=0):
    """Extrahiert"""
    text = ""

    try:
        # OCR-Verarbeitung mit OCRmyPDF und Speichern des Texts im Speicher
        temp_txt_file = pdf_path.replace('.pdf', '_text.txt')
        ocrmypdf.ocr(pdf_path, temp_txt_file, force_ocr=True, sidecar=temp_txt_file)
        with open(temp_txt_file, 'r') as txt_file:
            text = txt_file.read()
        os.remove(temp_txt_file)

        log_message(f"Text erfolgreich extrahiert für Datei {pdf_path}")
    except Exception as e:
        if retry_count < 3:
            log_message(f"Fehler bei der Textverarbeitung von {pdf_path}: {e}. Wiederhole ({retry_count + 1}/3)")
            time.sleep(5)
            return ocr_and_extract_text(pdf_path, retry_count=retry_count + 1)
        else:
            log_message(f"Fehler bei der Textverarbeitung von {pdf_path} nach mehreren Versuchen: {e}")
            return ""

    return text


# suche fettgedrucktes und Schlüsselwörter
def get_subject_from_text(text):
    """Sucht"""
    subject = None

    try:
        lines = text.split("\n")

        # Suche nach bestimmten Schlüsselwörtern
        for line in lines:
            if re.search(r'akten-\s*geschäftszeichen[:\s]*[a-zA-Z0-9\/\-\.]+', line, re.IGNORECASE):
                # Extrahiere das Aktenzeichen, das unterhalb von "Akten- Geschäftszeichen" steht
                aktenzeichen_line_index = lines.index(line) + 1
                if aktenzeichen_line_index < len(lines):
                    next_line = lines[aktenzeichen_line_index].strip()
                    date_match = re.search(r'\d{2}\.\d{2}\.\d{4}', next_line)
                    if date_match:
                        extracted_string = next_line.split(date_match.group())[0].strip()
                        log_message(f"String09: {extracted_string}")
                        subject = f"Aktenzeichen_{extracted_string}"
                break
            elif re.search(r'gericht|beschluss|urteil', line, re.IGNORECASE):
                subject = line.strip()
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

    if not os.path.exists(new_file_path):
        os.rename(old_path, new_file_path)
        log_message(f"Datei umbenannt von {old_file_name} zu {new_file_name}")
    else:
        log_message(f"Fehler: Die Datei {new_file_path} existiert bereits.")


# Dateinamen bereinigen
def sanitize_filename(filename):
    """Entfernt"""
    return re.sub(r'[\/\:*?"<>|]', '_', filename)


# Hauptprozess
def process_scan_files(scans_folder):
    """Durchsucht"""
    log_message(f"Arbeitsverzeichnis: {os.getcwd()}")
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf") and not file_name.startswith("._"):
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            # OCR-Verarbeitung und Text Extraktion
            text = ocr_and_extract_text(file_path)
            if not text:
                continue

            subject = get_subject_from_text(text)

            if subject:
                absender = file_name.split('_')[1]  # Absender an zweiter Stelle im Dateinamen
                date = file_name.split('_')[0]  # Datum an erster Stelle im Dateinamen
                # Sicherstellen, dass das Datum im korrekten Format ist
                if re.match(r'\d{8}', date):
                    rename_file(file_path, absender, date, subject)
                else:
                    log_message(f"Fehler: Datum im falschen Format für {file_name}")
            else:
                log_message(f"Kein Betreff gefunden für {file_name}.")

            # Fehlerhafte Dateien überspringen, um Loops zu verhindern
            if "._" in file_name:
                log_message(f"Überspringe Datei {file_name} aufgrund eines möglichen Fehlers im PDF-Format.")
                continue


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
