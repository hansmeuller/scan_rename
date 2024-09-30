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


#logfile prüfen/create
def check_and_create_log_file():
    """create if not exist"""
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")


#einträge mit timestamp
def log_message(message):
    """logdatei + zeitstempel"""
    check_and_create_log_file()  # create if not
    clean_old_log_entries()  # lösche+=woche

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


#nur einträge -=woche
def clean_old_log_entries():
    """create if not exist"""
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            lines = log_file.readlines()

        #achte auf speicher->nicht älter als 7 tage
        woche_alt = datetime.now() - timedelta(days=7)
        with open(log_file_path, "w") as log_file:
            for line in lines:
                try:
                    log_time = datetime.strptime(line.split(" - ")[0], "%Y-%m-%d %H:%M:%S")
                    if log_time > woche_alt:
                        log_file.write(line) #nix behalten>woche
                except (ValueError, IndexError):
                    log_file.write(line)


def extract_text_from_image(image_path):
    """extrahiere aus Bild"""
    try:
        with Image.open(image_path) as img:
            return pytesseract.image_to_string(img)
    except Exception as e:
        print(f"Erreur de traitement de l'image {image_path}: {e}")
        return ""



#apple autostart
def create_launch_agent():
    """create if not exist"""
    launch_agent_path = os.path.expanduser("~/Library/LaunchAgents/com.meinprogramm.scanrenamer.plist")

    if not os.path.exists(launch_agent_path):
        program_path = os.path.realpath(__file__)

        plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
        <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
        <plist version="1.0">
        <dict>
            <key>Label</key>
            <string>com.meinprogramm.scanrenamer</string>

            <key>ProgramArguments</key>
            <array>
                <string>{program_path}</string>
            </array>

            <key>RunAtLoad</key>
            <true/>

            <key>KeepAlive</key>
            <true/>
        </dict>
        </plist>"""

        with open(launch_agent_path, "w") as plist_file:
            plist_file.write(plist_content)

        os.system(f"launchctl load {launch_agent_path}")
        log_message("LaunchAgent wurde erstellt und geladen.")
    else:
        log_message("LaunchAgent existiert bereits.")


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
        log_message(f"Datei umbenannt von {old_file_name} zu {new_name}.{file_extension}")
    else:
        log_message(f"Fehler: Die Datei {new_file_path} existiert bereits.")


def process_scan_files(scans_folder):
    """bearbeite inkrementell"""
    for file_name in os.listdir(scans_folder):
        if file_name.endswith((".png", ".jpg", ".pdf")) and "Anzahl" in file_name:
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            text = extract_text_from_image(file_path)
            subject = get_subject_from_text(text)

            if subject:
                new_name = file_name.replace("Anzahl", subject)  # Ersetze Anzahl durch den Betreff
                rename_file(file_path, new_name)
            else:
                log_message(f"Kein Betreff gefunden für {file_name}.")


def main():
    """hier wird ausgeführt"""
    try:
        process_scan_files(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")
        time.sleep(5)  # warten
        main()  # Programm restart


if __name__ == "__main__":
    create_launch_agent()
    log_message("Programm gestartet.")
    while True:
        try:
            main()  # Hauptverarbeitung
        except Exception as e:
            log_message(f"Fehler im Hauptprogramm: {e}")
            time.sleep(5)  # warte restart
