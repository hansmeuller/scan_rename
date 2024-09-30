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


# Logdatei prüfen/create
def check_and_create_log_file():
    """create if not exist"""
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")


# Einträge mit timestamp
def log_message(message):
    """logdatei + zeitstempel"""
    check_and_create_log_file()  # create if not
    clean_old_log_entries()  # lösche+=woche

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# Nur Einträge -=woche
def clean_old_log_entries():
    """Lösche Einträge älter als eine Woche"""
    if os.path.exists(log_file_path):
        with open(log_file_path, "r") as log_file:
            lines = log_file.readlines()

        # Achte auf speicher → nicht älter als 7 Tage
        woche_alt = datetime.now() - timedelta(days=7)
        with open(log_file_path, "w") as log_file:
            for line in lines:
                try:
                    log_time = datetime.strptime(line.split(" - ")[0], "%Y-%m-%d %H:%M:%S")
                    if log_time > woche_alt:
                        log_file.write(line)  # nichts behalten > Woche
                except (ValueError, IndexError):
                    log_file.write(line)


def extract_text_from_image(image_path):
    """extrahiere aus Bild"""
    try:
        with Image.open(image_path) as img:
            return pytesseract.image_to_string(img)
    except Exception as e:
        log_message(f"Erreur de traitement de l'image {image_path}: {e}")
        return ""


# Apple Autostart
def check_and_update_launch_agent():
    """Prüft LaunchAgent-Pfad und aktualisiert, falls nötig."""
    launch_agent_path = os.path.expanduser("~/Library/LaunchAgents/com.meinprogramm.scanrenamer.plist")

    # Dynamischer Pfad
    current_program_path = os.path.realpath(__file__)

    # check if exist
    if os.path.exists(launch_agent_path):
        # Inhalt lesen
        with open(launch_agent_path, "r") as plist_file:
            content = plist_file.read()

        # Pfad == LaunchAgent?
        if current_program_path in content:
            log_message(f"LaunchAgent ist aktuell. Programm wird von {current_program_path} ausgeführt.")
        else:
            log_message("LaunchAgent veraltet. Wird aktualisiert.")
            create_launch_agent(current_program_path)
    else:
        log_message("LaunchAgent existiert nicht. Erstellt neuen LaunchAgent.")
        create_launch_agent(current_program_path)


def create_launch_agent():
    """create if not exist+dynamische pfad erkennung"""
    launch_agent_path = os.path.expanduser("~/Library/LaunchAgents/com.meinprogramm.scanrenamer.plist")

    # Dynamische Erkennung des aktuellen Programmpfads
    program_path = os.path.realpath(__file__)

    plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
    <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
    <plist version="1.0">
    <dict>
        <key>Label</key>
        <string>com.meinprogramm.scanrenamer</string>

        <key>ProgramArguments</key>
        <array>
            <string>/usr/bin/python3</string>  <!-- Python Interpreter -->
            <string>{program_path}</string>    <!-- Dynamischer Pfad zum Programm -->
        </array>

        <key>RunAtLoad</key>
        <true/>

        <key>KeepAlive</key>
        <true/>
    </dict>
    </plist>"""

    # LaunchAgent Datei erstellen
    with open(launch_agent_path, "w") as plist_file:
        plist_file.write(plist_content)

    # LaunchAgent laden
    os.system(f"launchctl load {launch_agent_path}")
    log_message(f"LaunchAgent wurde erstellt und geladen für {program_path}.")


def get_subject_from_text(text):
    """fettgedruckt oder Begriffe"""
    subject = None
    lines = text.split('\n')

    # Such "Akten-/Geschäftszeichen" → nimm die Zeile darunter
    for i, line in enumerate(lines):
        if "Akten-/Geschäftszeichen" in line:
            try:
                subject = f"Aktenzeichen_{lines[i + 1].strip()}"  # nächste Zeile als Betreff
                break
            except IndexError:
                log_message("Kein Aktenzeichen in der Zeile gefunden.")
                break

    # falls kein Betreff
    if not subject:
        for line in lines:
            if "Aktenzeichen" in line or "Geschäftszeichen" in line:
                subject = line.strip()
                break
    return subject


def rename_file(old_path, absender, date, subject):
    """benenne um → Absender_Datum_Betreff"""
    directory, old_file_name = os.path.split(old_path)
    file_extension = old_file_name.split('.')[-1]
    new_file_name = f"{absender}_{date}_{subject}.{file_extension}"  # ANFORDERUNG-> Absender_Datum_Betreff
    new_file_path = os.path.join(directory, new_file_name)

    if not os.path.exists(new_file_path):  # vermeide Konflikte
        os.rename(old_path, new_file_path)
        log_message(f"Datei umbenannt von {old_file_name} zu {new_file_name}")
    else:
        log_message(f"Fehler: Die Datei {new_file_path} existiert bereits.")


def process_scan_files(scans_folder):
    """bearbeite inkrementell"""
    for file_name in os.listdir(scans_folder):
        if file_name.endswith((".png", ".jpg", ".pdf")) and "Anzahl" in file_name:
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            absender = file_name.split('_')[1]  # annahme Absender an zweiter Stelle
            date = file_name.split('_')[0]  # Datum an erster

            text = extract_text_from_image(file_path)
            subject = get_subject_from_text(text)

            if subject:
                rename_file(file_path, absender, date, subject)  # Übergabe -> an rename_file
            else:
                log_message(f"Kein Betreff gefunden für {file_name}.")


def main():
    """Hauptverarbeitung"""
    try:
        process_scan_files(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")
        time.sleep(5)  # Warten
        main()  # Programm neu starten


if __name__ == "__main__":
    create_launch_agent()
    log_message("Programm gestartet.")
    while True:
        try:
            main()  # Hauptverarbeitung
        except Exception as e:
            log_message(f"Fehler im Hauptprogramm: {e}")
            time.sleep(5)  # Warten vor Neustart
