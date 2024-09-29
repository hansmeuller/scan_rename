import time
import os
from PIL import Image
import pytesseract
import re
from datetime import datetime, timedelta


# scant im verzeichnis
scans_folder = os.path.dirname(os.path.realpath(__file__))


# Pfad Logdatei
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
log_file_path = os.path.join(desktop_path, "Logeinträge_ScanRename.txt")

def check_and_create_log_file():
    """Logdatei erstellt falls notwendig"""
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")


def extract_text_from_image(image_path):
    """extrahiere aus Bild"""
    try:
        with Image.open(image_path) as img:
            return pytesseract.image_to_string(img)
    except Exception as e:
        print(f"Erreur de traitement de l'image {image_path}: {e}")
        return ""

        # Nur Einträge behalten, die jünger als 7 Tage sind
        one_week_ago = datetime.now() - timedelta(days=7)
        with open(log_file_path, "w") as log_file:
            for line in lines:
                # Extrahiere das Datum aus dem Zeitstempel (angenommen Format: YYYY-MM-DD HH:MM:SS)
                try:
                    log_time = datetime.strptime(line.split(" - ")[0], "%Y-%m-%d %H:%M:%S")
                    if log_time > one_week_ago:
                        log_file.write(line)  # Nur Einträge behalten, die neuer als eine Woche sind
                except (ValueError, IndexError):
                    log_file.write(line)  # Falls das Format nicht stimmt, trotzdem die Zeile behalten


def log_message(message):
    """Schreibt in Logdatei mit Zeitstempel."""
    check_and_create_log_file()  # Prüfen, ob die Logdatei existiert
    clean_old_log_entries()  # Alte Einträge bereinigen

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


def get_subject_from_text(text):
    """suche fettgedrucktes"""
    subject = None
    lines = text.split('\n')

    # nur erstes Drittel
    for line in lines[:15]:
        if line.isupper():  # Großbuchstaben simulieren
            subject = line.strip()
            break

    # Falls kein Betreff
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

    if not os.path.exists(new_file_path):  # vermeidung Konflikte
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


def main():
    """hier wird ausgeführt"""
    try:
        process_scan_files(scans_folder)
    except Exception as e:
        print(f"Programm abgestürzt: {e}")
        time.sleep(5)  # warten
        main()  # Programm neu starten


if __name__ == "__main__":
    while True:
        try:
            main()  # starte Hauptverarbeitung
        except Exception as e:
            print(f"Fehler im Hauptprogramm: {e}")
            time.sleep(5)  # warte restart
