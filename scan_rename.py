import os
import fitz  # PyMuPDF
import re
from datetime import datetime

# Pfad zum Ordner mit den PDFs
scans_folder = os.path.dirname(os.path.realpath(__file__))
log_file_path = os.path.join(scans_folder, "Logeinträge_TestExtraction.txt")

# Knickfalte
KNICKFALTE_Y_RATIO = 1 / 3
KONTOAUSZUG_MAX_HEIGHT = 600  # Kontoauszüge

# Log
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")

# Extraktion
def extract_text_from_fold(pdf_path):
    try:
        with fitz.open(pdf_path) as pdf_document:
            for page_number in range(len(pdf_document)):
                page = pdf_document.load_page(page_number)
                page_height = page.rect.height
                knickfalte_y = page_height * KNICKFALTE_Y_RATIO

                # Extrahiere
                words = page.get_text("words")
                full_text = page.get_text("text")

                # Kontoauszug
                bank_keywords = ["kontoauszug", "kontostand", "privatkonto", "blz"]
                if page_height < KONTOAUSZUG_MAX_HEIGHT:
                    log_message(f"Gefundener Betreff für {pdf_path}: Kontoauszug")
                    return
                elif any(keyword in full_text.lower() for keyword in bank_keywords):
                    log_message(f"Gefundener Betreff für {pdf_path}: Kontoauszug")
                    return

                # Aktenzeichen extrahieren
                marker_pattern = re.compile(r"bitte bei antwort angeben", re.IGNORECASE)
                marker_match = marker_pattern.search(full_text)
                if marker_match:
                    marker_line = full_text.splitlines()
                    marker_index = [i for i, line in enumerate(marker_line) if marker_pattern.search(line)]
                    if marker_index:
                        target_line_index = marker_index[0] + 2  # Zwei Zeilen unter
                        if target_line_index < len(marker_line):
                            aktenzeichen_text = marker_line[target_line_index].strip()[:15]
                            log_message(f"Gefundenes Aktenzeichen für {pdf_path}: {aktenzeichen_text}")
                            return  # Aktenzeichen gefunden

                # unterhalb Knickfalte extrahieren
                knick_text = []
                word_count = 0
                for w in words:
                    if w[1] >= knickfalte_y:
                        knick_text.append(w[4])
                        word_count += 1
                        if word_count >= 5 or '.' in w[4] or '?' in w[4] or '!' in w[4]:  # Stop 5 Wörter
                            break
                log_message(f"Erster Text unterhalb der Knickfalte in {pdf_path}: {' '.join(knick_text)}")
                return

    except Exception as e:
        log_message(f"Fehler beim Extrahieren des Textes aus {pdf_path}: {e}")

# Hauptprozess
def process_pdfs(scans_folder):
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf") and not file_name.startswith("._"):
            file_path = os.path.join(scans_folder, file_name)
            extract_text_from_fold(file_path)

# Hauptverarbeitung
def main():
    try:
        log_message("Start erfolgreich")
        process_pdfs(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")

if __name__ == "__main__":
    main()