from PIL import Image
import pytesseract
import os
import pdf2image
import re

# scant im verzeichnis in dem python file liegt
scans_folder = os.path.dirname(os.path.realpath(__file__))

def process_scan_files(scans_folder):
    """scan verzeichnis und bearbeite """
    for file_name in os.listdir(scans_folder):
        if file_name.endswith((".png", ".jpg", ".pdf")):
            file_path = os.path.join(scans_folder, file_name)
            print(f"verarbeite {file_name}...")


def extract_text_from_image(image_path):
    """extrahiere text"""
    try:
        with Image.open(image_path) as img:
            return pytesseract.image_to_string(img)
    except Exception as e:
        print(f"Erreur de traitement de l'image {image_path}: {e}")
        return ""


def get_subject_from_text(text):
    """suche nach fettgedrucktem"""
    subject = None
    lines = text.split('\n')
    for line in lines[:15]:  # nur erste drittel der Seite analysieren
        if line.isupper():  # bsp fettgedruckt simuliert durch großbuchstaben
            subject = line.strip()
            break
    if not subject:
        # suche nach words
        for line in lines:
            if "Aktenzeichen" in line or "Geschäftszeichen" in line:
                subject = line.strip()
                break
    return subject


def rename_file(old_path, new_name):
    """Benenne um"""
    directory, old_file_name = os.path.split(old_path)
    file_extension = old_file_name.split('.')[-1]
    new_file_path = os.path.join(directory, f"{new_name}.{file_extension}")
    os.rename(old_path, new_file_path)
    print(f"Datei umbenannt von {old_file_name} zu {new_name}.{file_extension}")


def process_scan_files(scans_folder):
    """durchsuche verzeichnis, bearbeite inkrementell."""
    for file_name in os.listdir(scans_folder):
        if file_name.endswith((".png", ".jpg", ".pdf")):
            file_path = os.path.join(scans_folder, file_name)

            # keine dateinamen mit Betreff
            if not "_Betreff_" in file_name:
                print(f"Verarbeite {file_name}...")
                text = extract_text_from_image(file_path)
                subject = get_subject_from_text(text)

                if subject:
                    new_name = file_name.replace("Anzahl", subject)  # Ersetze Anzahl durch den Betreff
                    rename_file(file_path, new_name)
                else:
                    print(f"Kein Betreff gefunden für {file_name}.")


if __name__ == "__main__":
    # Starte die Verarbeitung
    process_scan_files(scans_folder)