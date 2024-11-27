import os
import easyocr
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
from datetime import datetime
import re

reader = easyocr.Reader(['de'], gpu=False)

WINDOW_LEFT_CM = 2
WINDOW_TOP_CM = 4.5
WINDOW_WIDTH_CM = 9
WINDOW_HEIGHT_CM = 4.5

KNICKFALTE_TOP_CM = 10
KNICKFALTE_HEIGHT_CM = 2

AKTENZEICHEN_KEYWORDS = ["Bitte bei"]
AUSSCHLUSSLISTE = ["Postfach", "PLZ", "Postzentrum"]
ZAHLEN_REGEX = r"\b\d+\b"
OCR_CORRECTIONS = {"unaedeckte": "ungedeckte"}


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


def normalize_spacing(text):
    return re.sub(r"\s{2,}", " ", text).strip()


def rename_file(original_path, sender, date, subject):
    try:
        folder = os.path.dirname(original_path)
        sender = normalize_spacing(sender.replace(" ", "_"))
        subject = normalize_spacing(subject.replace(" ", "_"))

        new_name = f"{sender}_{date}_{subject}.pdf"
        new_path = os.path.join(folder, new_name)

        if original_path != new_path:
            os.rename(original_path, new_path)
            log_message(f"Datei umbenannt: {original_path} -> {new_path}")
        else:
            log_message(f"Keine Umbenennung erforderlich: {original_path}")

    except FileNotFoundError:
        log_message(f"Fehler beim Umbenennen der Datei {original_path}: Datei nicht gefunden.")
    except Exception as e:
        log_message(f"Fehler beim Umbenennen der Datei {original_path}: {e}")


def process_pdf(file_path):
    try:
        file_name = os.path.basename(file_path)

        match = re.search(r"\b\d{8}\b", file_name)
        if not match:
            log_message(f"Kein gültiges Datum im Dateinamen gefunden: {file_name}")
            return
        date = match.group()

        images = convert_from_path(file_path, dpi=300, first_page=1, last_page=1)
        if not images:
            log_message(f"Keine Bilder extrahiert für {file_path}")
            return

        image = images[0]
        dpi = 300

        sender_results = extract_text_from_window(image, dpi, file_name, top_cm=WINDOW_TOP_CM,
                                                  height_cm=WINDOW_HEIGHT_CM)
        sender = extract_sender(sender_results)

        subject_or_case = extract_subject_or_case_number(image, dpi, file_name)

        rename_file(file_path, sender, date, subject_or_case)

    except Exception as e:
        log_message(f"Fehler bei der Verarbeitung von {file_path}: {e}")


def process_pdfs(folder):
    for file_name in os.listdir(folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder, file_name)
            log_message(f"Starte Verarbeitung für Datei: {file_path}")
            process_pdf(file_path)


def extract_text_from_window(image, dpi, file_name, top_cm, height_cm):
    try:
        pixels_per_cm = dpi / 2.54
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
        first_line = re.sub(ZAHLEN_REGEX, "", first_line)
        first_line = normalize_spacing(first_line)
        return first_line
    return "Unbekannter_Absender"


def extract_case_number(results):
    for idx, result in enumerate(results):
        line_text = normalize_spacing(result[1].strip())

        if any(line_text.startswith(keyword) for keyword in AKTENZEICHEN_KEYWORDS):
            log_message(f"Keyword für Aktenzeichen gefunden: {line_text}")

            if idx + 3 < len(results):
                aktenzeichen_line = normalize_spacing(results[idx + 3][1].strip())
                return f"Aktenzeichen: {aktenzeichen_line[:11]}"

    return None


def extract_subject_or_case_number(image, dpi, file_name):
    results = extract_text_from_window(image, dpi, file_name, top_cm=KNICKFALTE_TOP_CM, height_cm=KNICKFALTE_HEIGHT_CM)
    if not results:
        return "Kein_Betreff"

    for idx, result in enumerate(results):
        line_text = normalize_spacing(result[1].strip())

        for key, correction in OCR_CORRECTIONS.items():
            line_text = line_text.replace(key, correction)

        case_number = extract_case_number(results)
        if case_number:
            return case_number

        if idx == 0:
            return f"{' '.join(line_text.split()[:5])}"  # Maximal 5 Worte

    return "Kein_Betreff"


if __name__ == "__main__":
    process_pdfs(os.getcwd())
