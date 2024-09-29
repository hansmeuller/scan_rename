from PIL import Image
import pytesseract
import os
import re

# scant im verzeichnis, in dem das Programm liegt
scans_folder = os.path.dirname(os.path.realpath(__file__))


def extract_text_from_image(image_path):
    """extrahiere aus Bild"""
    try:
        with Image.open(image_path) as img:
            return pytesseract.image_to_string(img)
    except Exception as e:
        print(f"Erreur de traitement de l'image {image_path}: {e}")
        return ""


def get_subject_from_text(text):
    """suche nach fettgedrucktem oder Begriffen"""
    subject = None
    lines = text.split('\n')

    # Nur das erste Drittel analysieren
    for line in lines[:15]:
        if line.isupper():  # Großbuchstaben simulieren fettgedruckt
            subject = line.strip()
            break

    # falls kein Betreff
    if not subject:
        for line in lines:
            if "Aktenzeichen" in line or "Geschäftszeichen" in line:
                subject = line.strip()
                break
    return subject


def rename_file(old_path, new_name):
    """benenne um"""
    directory, old_file_name = os.path.split(old_path)
    file_extension = old_file_name.split('.')[-1]
    new_file_path = os.path.join(directory, f"{new_name}.{file_extension}")

    if not os.path.exists(new_file_path):  # Vermeidung von Konflikten mit bestehenden Dateinamen
        os.rename(old_path, new_file_path)
        print(f"Datei umbenannt von {old_file_name} zu {new_name}.{file_extension}")
    else:
        print(f"Fehler: Die Datei {new_file_path} existiert bereits.")


def process_scan_files(scans_folder):
    """durchsuche und bearbeite inkrementell"""
    for file_name in os.listdir(scans_folder):
        if file_name.endswith((".png", ".jpg", ".pdf")) and "Anzahl" in file_name:
            file_path = os.path.join(scans_folder, file_name)
            print(f"Verarbeite {file_name}...")

            text = extract_text_from_image(file_path)
            subject = get_subject_from_text(text)

            if subject:
                new_name = file_name.replace("Anzahl", subject)  # Ersetze Anzahl durch den Betreff
                rename_file(file_path, new_name)
            else:
                print(f"Kein Betreff gefunden für {file_name}.")


if __name__ == "__main__":
    # starte die Verarbeitung
    process_scan_files(scans_folder)
