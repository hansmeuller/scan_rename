import time
import os
import re
import ocrmypdf
import fitz  # PyMuPDF
from datetime import datetime, timedelta

# current
scans_folder = os.path.dirname(os.path.realpath(__file__))
log_file_path = os.path.join(scans_folder, "Logeinträge_ScanRename.txt")

# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")

# kill +=woche
def clean_old_log_entries():
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

# Prüfen
def is_correctly_named(file_name):
    # Absender_Datum_Betreff
    pattern = r"^[^_]+_\d{8}_.+\.pdf$"
    return re.match(pattern, file_name) is not None

# OCR
def ocr_and_extract_text(pdf_path):
    temp_pdf_path = pdf_path.replace('.pdf', '_ocr.pdf')
    ocrmypdf.ocr(pdf_path, temp_pdf_path, force_ocr=True, output_type="pdfa")
    return temp_pdf_path

# Extrahiere
def extract_bold_text_and_key_elements(pdf_path):
    bold_text = []
    aktenzeichen = None
    kontoauszug = False
    try:
        with fitz.open(pdf_path) as pdf_document:
            for page_number in range(len(pdf_document)):
                page = pdf_document.load_page(page_number)
                page_height = page.rect.height
                blocks = page.get_text("dict")['blocks']
                for block in blocks:
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            if span.get("flags", 0) & 2:  # Fettgedruckter Text
                                bold_text.append(span.get("text", "").strip())
                            elif re.search(r'akten.?zeichen|geschäfts.?zeichen', span.get("text", ""), re.IGNORECASE):
                                aktenzeichen = next(
                                    (l.get("spans", [])[0].get("text") for l in block["lines"] if len(l["spans"]) > 1), None)
                            elif re.search(r'sparkasse|bank|konto', span.get("text", ""), re.IGNORECASE):
                                kontoauszug = True
    except Exception as e:
        log_message(f"Fehler beim Extrahieren aus {pdf_path}: {e}")
    return " ".join(bold_text), aktenzeichen, kontoauszug

# priorisieren
def get_subject(bold_text, aktenzeichen, kontoauszug):
    if bold_text:
        return bold_text[:30]  # 30 Zeichen
    if aktenzeichen:
        return f"Aktenzeichen {aktenzeichen}"
    if kontoauszug:
        return "Kontoauszug"
    return "Betreff unbekannt"

# umbenennen
def rename_file(old_path, absender, date, subject):
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
    return re.sub(r'[\/\:*?"<>|]', '_', filename)

# Hauptprozess
def process_scan_files(scans_folder):
    log_message(f"Arbeitsverzeichnis: {os.getcwd()}")
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf") and not file_name.startswith("._"):
            if is_correctly_named(file_name):
                log_message(f"{file_name} entspricht bereits den Anforderungen. Datei wird übersprungen.")
                continue  # skipp leg day

            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            absender = file_name.split('_')[1] if len(file_name.split('_')) > 1 else "Unbekannt"
            date_match = re.match(r'\d{8}', file_name)
            date = date_match.group(0) if date_match else datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y%m%d')

            temp_pdf = ocr_and_extract_text(file_path)
            bold_text, aktenzeichen, kontoauszug = extract_bold_text_and_key_elements(temp_pdf)
            subject = get_subject(bold_text, aktenzeichen, kontoauszug)

            rename_file(file_path, absender, date, subject)
            os.remove(temp_pdf)  # kill temp

# main
def main():
    try:
        log_message("Start erfolgreich")
        clean_old_log_entries()
        process_scan_files(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")
        time.sleep(5)
        main()

if __name__ == "__main__":
    main()
