import os
import easyocr
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
from datetime import datetime

reader = easyocr.Reader(['de'], gpu=False)

WINDOW_LEFT_CM = 2
WINDOW_TOP_CM = 4.5
WINDOW_WIDTH_CM = 9
WINDOW_HEIGHT_CM = 4.5

KNICKFALTE_TOP_CM = 10
KNICKFALTE_HEIGHT_CM = 2

AKTENZEICHEN_KEYWORDS = ["Bitte bei Antwort angeben"]
AUSSCHLUSSLISTE = ["Postfach", "PLZ", "12345", "45678"]


def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("Logeinträge.txt", "a", encoding="utf-8") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


def delete_temp_png(file_path):
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            log_message(f"Temporäre Datei gelöscht: {file_path}")
    except Exception as e:
        log_message(f"Fehler beim Löschen der PNG: {e}")


def extract_text_from_window(image, dpi, file_name, top_cm, height_cm):
    try:
        pixels_per_cm = dpi / 2.54  # 1 Zoll = 2.54 cm
        left = int(WINDOW_LEFT_CM * pixels_per_cm)
        right = int((WINDOW_LEFT_CM + WINDOW_WIDTH_CM) * pixels_per_cm)
        top = int(top_cm * pixels_per_cm)
        bottom = int((top_cm + height_cm) * pixels_per_cm)

        cropped_image = image.crop((left, top, right, bottom))
        window_image_path = f"{file_name}_window_preview.png"
        cropped_image.save(window_image_path)
        log_message(f"Fensterbereich gespeichert: {window_image_path}")

        cropped_array = np.array(cropped_image)
        results = reader.readtext(cropped_array)
        log_message(f"OCR-Ergebnisse im Fenster: {results}")

        delete_temp_png(window_image_path)

        return results
    except Exception as e:
        log_message(f"Fehler bei der Textauswertung: {e}")
        return []


def extract_sender(results):
    if results:
        first_line = results[0][1]
        for exclusion in AUSSCHLUSSLISTE:
            first_line = first_line.replace(exclusion, "")
        for delimiter in [",", ";", "."]:
            if delimiter in first_line:
                first_line = first_line.split(delimiter)[0]
                break
        return f"Absender: {first_line.strip()}"
    return "Kein Absender gefunden"


def extract_subject_or_case_number(image, dpi, file_name):
    results = extract_text_from_window(image, dpi, file_name, KNICKFALTE_TOP_CM, KNICKFALTE_HEIGHT_CM)
    if not results:
        return "Kein Betreff gefunden"

    for result in results:
        line_text = result[1].strip()

        if any(keyword in line_text for keyword in AKTENZEICHEN_KEYWORDS):
            log_message(f"Keyword für Aktenzeichen gefunden: {line_text}")
            idx = results.index(result)
            if idx + 2 < len(results):
                aktenzeichen_line = results[idx + 2][1].strip()
                return f"Aktenzeichen: {aktenzeichen_line[:10]}"
            return "Kein Aktenzeichen gefunden"

        return f"Betreff: {' '.join(line_text.split()[:5])}"


def is_kontoauszug(image):
    if image.height < 1500:
        return True
    return False


def process_pdf(file_path):
    try:
        file_name = os.path.basename(file_path)
        images = convert_from_path(file_path, dpi=300, first_page=1, last_page=1)
        if not images:
            log_message(f"Keine Bilder extrahiert für {file_path}")
            return

        dpi = 300
        image = images[0]

        if is_kontoauszug(image):
            log_message(f"Dokument erkannt als Kontoauszug: {file_path}")
            log_message(f"Betreff: Kontoauszug")
            sender_results = extract_text_from_window(image, dpi, file_name, WINDOW_TOP_CM, WINDOW_HEIGHT_CM)
            sender = extract_sender(sender_results)
            log_message(f"Gefundener Absender in {file_path}: {sender}")
            return

        sender_results = extract_text_from_window(image, dpi, file_name, WINDOW_TOP_CM, WINDOW_HEIGHT_CM)
        sender = extract_sender(sender_results)
        log_message(f"Gefundener Absender in {file_path}: {sender}")

        subject_or_case = extract_subject_or_case_number(image, dpi, file_name)
        log_message(f"Gefundener Betreff oder Aktenzeichen in {file_path}: {subject_or_case}")

    except Exception as e:
        log_message(f"Fehler bei der Verarbeitung von {file_path}: {e}")


def process_pdfs(folder):
    for file_name in os.listdir(folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder, file_name)
            log_message(f"Starte Verarbeitung für Datei: {file_path}")
            process_pdf(file_path)


if __name__ == "__main__":
    process_pdfs(os.getcwd())
