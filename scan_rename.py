import os
import easyocr
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
from datetime import datetime
import re
from PyPDF2 import PdfReader

# init
reader = easyocr.Reader(['de'], gpu=False)

# fenster
WINDOW_LEFT_CM = 2
WINDOW_TOP_CM = 4.5
WINDOW_WIDTH_CM = 9
WINDOW_HEIGHT_CM = 4.5

# Knickfalte
KNICKFALTE_TOP_CM = 10
KNICKFALTE_HEIGHT_CM = 2

# Keywords
AKTENZEICHEN_KEYWORDS = ["Bitte bei"]
AUSSCHLUSSLISTE = ["Postfach", "PLZ", "Postzentrum"]
ZAHLEN_REGEX = r"\b\d+\b"
OCR_CORRECTIONS = {"unaedeckte": "ungedeckte"}


# lognachrichten mit zeitstempel
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("Logeinträge.txt", "a", encoding="utf-8") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# png löschen
def delete_temp_png(file_path):
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            log_message(f"Temporäre Datei gelöscht: {file_path}")
    except Exception as e:
        log_message(f"Fehler beim Löschen der PNG: {e}")


# max leerzeichen
def normalize_spacing(text):
    return re.sub(r"\s{2,}", " ", text).strip()


# entfernen
def sanitize_filename(filename):
    return re.sub(r"[\\/*?\"<>|]", "", filename)


# unicorn generieren
def get_unique_filename(file_path):
    """
    check for another
    """
    base, ext = os.path.splitext(file_path)
    counter = 1

    # if dateiname füge eine Nummer hinzu
    while os.path.exists(file_path):
        file_path = f"{base} ({counter}){ext}"
        counter += 1

    return file_path


# aus fenster extrahieren
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


# abs extrahieren
def extract_sender(results):
    if results:
        first_line = results[0][1]
        for exclusion in AUSSCHLUSSLISTE:
            first_line = first_line.replace(exclusion, "")
        for delimiter in [",", ";", "."]:
            if delimiter in first_line:
                first_line = first_line.split(delimiter)[0]
                break
        first_line = re.sub(ZAHLEN_REGEX, "", first_line)  # zahlen entfernen
        first_line = normalize_spacing(first_line)  # max leerzeichen
        return first_line
    return "Unbekannter Absender"


# az gezielt extrahieren
def extract_case_number(results):
    extracted_case_numbers = []
    for idx, result in enumerate(results):
        line_text = normalize_spacing(result[1].strip())

        # extraktion einleiten kw
        if any(line_text.startswith(keyword) for keyword in AKTENZEICHEN_KEYWORDS):
            log_message(f"Keyword für Aktenzeichen gefunden: {line_text}")

            # 2 zeilen
            if idx + 3 < len(results):  # 3 zeilen
                aktenzeichen_line = normalize_spacing(results[idx + 3][1].strip())
                aktenzeichen_line = re.sub(r"[-/\\]$", "", aktenzeichen_line).strip()
                if len(aktenzeichen_line) > 1:
                    extracted_case_numbers.append(aktenzeichen_line[:13])  # max 13 Zeichen

    if extracted_case_numbers:
        log_message(f"Alle extrahierten Aktenzeichen: {extracted_case_numbers}")
        return max(extracted_case_numbers, key=len)  # längste az auswählen

    return None


# extrahieren
def extract_subject_or_case_number(image, dpi, file_name):
    results = extract_text_from_window(image, dpi, file_name, KNICKFALTE_TOP_CM, KNICKFALTE_HEIGHT_CM)
    if not results:
        return "Kein Betreff"

    for idx, result in enumerate(results):
        line_text = normalize_spacing(result[1].strip())

        # ocr-Korrektur
        for key, correction in OCR_CORRECTIONS.items():
            line_text = line_text.replace(key, correction)

        # az extrahieren, falls Keywords gefunden
        case_number = extract_case_number(results)
        if case_number:
            return case_number

        # extrahieren
        if idx == 0:
            return line_text  # Betreff mit Leerzeichen erhalten

    return "Kein Betreff"


# kto prüfen
def is_kontoauszug(image):
    if image.height < 1500:
        return True
    return False


# metadaten extrahieren -date
def extract_creation_date(file_path):
    try:
        reader = PdfReader(file_path)
        metadata = reader.metadata
        creation_date = metadata.get("/CreationDate", "")
        if creation_date:
            # Format anpassen: yyyymmdd
            match = re.search(r"D:(\d{4})(\d{2})(\d{2})", creation_date)
            if match:
                return f"{match.group(1)}{match.group(2)}{match.group(3)}"
        return "Unbekanntes Datum"
    except Exception as e:
        log_message(f"Fehler bei der Extraktion des Erstelldatums: {e}")
        return "Unbekanntes Datum"


# verarbeiten
def process_pdf(file_path):
    try:
        file_name = os.path.basename(file_path)
        images = convert_from_path(file_path, dpi=300, first_page=1, last_page=1)
        if not images:
            log_message(f"Keine Bilder extrahiert für {file_path}")
            return

        dpi = 300
        image = images[0]

        # kto erkennen
        if is_kontoauszug(image):
            log_message(f"Dokument erkannt als Kontoauszug: {file_path}")
            log_message(f"Betreff: Kontoauszug")
            sender_results = extract_text_from_window(image, dpi, file_name, WINDOW_TOP_CM, WINDOW_HEIGHT_CM)
            sender = extract_sender(sender_results)
            log_message(f"Gefundener Absender in {file_path}: {sender}")
            subject_or_case = "Kontoauszug"
        else:
            # abs extrahieren
            sender_results = extract_text_from_window(image, dpi, file_name, WINDOW_TOP_CM, WINDOW_HEIGHT_CM)
            sender = extract_sender(sender_results)
            log_message(f"Gefundener Absender in {file_path}: {sender}")

            # betreff az extrahieren
            subject_or_case = extract_subject_or_case_number(image, dpi, file_name)
            log_message(f"Gefundener Betreff oder Aktenzeichen in {file_path}: {subject_or_case}")

        # date extrahieren
        creation_date = extract_creation_date(file_path)
        log_message(f"Erstelldatum extrahiert: {creation_date}")

        # dateinamen erstellen
        new_file_name = f"{sender}_{creation_date}_{subject_or_case}.pdf"
        new_file_name = normalize_spacing(new_file_name).replace("__", "_")
        new_file_name = sanitize_filename(new_file_name)
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)

        # pink fluffy unicorn
        unique_file_path = get_unique_filename(new_file_path)

        # rename
        os.rename(file_path, unique_file_path)
        log_message(f"Datei umbenannt: {file_path} -> {unique_file_path}")

    except Exception as e:
        log_message(f"Fehler bei der Verarbeitung von {file_path}: {e}")


# current
def process_pdfs(folder):
    folder = os.getcwd()
    for file_name in os.listdir(folder):
        if file_name.lower().endswith(".pdf"):
            file_path = os.path.join(folder, file_name)
            process_pdf(file_path)



if __name__ == "__main__":
    folder_to_process = os.getcwd()  # current
    process_pdfs(folder_to_process)

# todo:

    #this in flutter
    #with gui