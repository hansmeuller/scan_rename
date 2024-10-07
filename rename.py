import time
import os
from PIL import Image
import re
from datetime import datetime, timedelta
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBoxHorizontal, LTTextLineHorizontal


# scant im verzeichnis
scans_folder = os.path.dirname(os.path.realpath(__file__))
# Pfad Logdatei
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
log_file_path = os.path.join(desktop_path, "Logeinträge_ScanRename.txt")
#pdf-pfad im dynamischen datei pfad
pdf_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "current_program_path.pdf")



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
    """Lösche +=Woche"""
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


def extract_text_with_format_from_pdf(pdf_path):
    """Extrahiert Text"""
    formatted_text = []

    for page_layout in extract_pages(pdf_path):
        for element in page_layout:
            if isinstance(element, LTTextBoxHorizontal) or isinstance(element, LTTextLineHorizontal):
                for text_line in element:
                    # Zugriff auf Textinhalt und mögliche Formatierungen
                    text = text_line.get_text()
                    font_size = text_line.height  # Annahme: Schriftgröße aus der Höhe des Texts
                    font_name = text_line.fontname if hasattr(text_line, 'fontname') else "unknown"

                    # Wir können bestimmte Kriterien nutzen, um Fettdruck zu simulieren
                    if font_name.lower().find("bold") != -1 or font_size > 12:  # Beispiel für große Schrift/Fettdruck
                        formatted_text.append((text.strip(), "bold"))
                    else:
                        formatted_text.append((text.strip(), "normal"))

    return formatted_text


formatted_text = extract_text_with_format_from_pdf(pdf_path)

# Ausgabe der Formatierung
for line, format in formatted_text:
    log_message(f"Text: {line}, Format: {format}")


def get_subject_from_formatted_text(formatted_text):
    """Such nach fettgedruckten"""
    subject = None

    for line, format in formatted_text:
        if format == "bold":  # Sucht nach fettgedruckten Zeilen
            subject = line.strip()
            break

    return subject

# Nutzung
formatted_text = extract_text_with_format_from_pdf(pdf_path)
subject = get_subject_from_formatted_text(formatted_text)


def process_scan_files(scans_folder):
    """bearbeite mit pdfminer"""
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf"):  # Jetzt nur PDFs verarbeiten
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            # ExtrahiereText mit pdfminer
            formatted_text = extract_text_with_format_from_pdf(file_path)
            subject = get_subject_from_formatted_text(formatted_text)

            if subject:
                new_name = file_name.replace("Anzahl", subject)  # Betreff in den Dateinamen einfügen
                rename_file(file_path, new_name)
            else:
                log_message(f"Kein Betreff gefunden für {file_name}.")


# Apple Autostart
def check_and_update_launch_agent():
    """LaunchAgent-Pfad +aktualisierung"""
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


def create_launch_agent(program_path):
    """nimmt angegebenen Programm-Pfad"""
    launch_agent_path = os.path.expanduser("~/Library/LaunchAgents/com.meinprogramm.scanrenamer.plist")

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

    # create or overwrite
    with open(launch_agent_path, "w") as plist_file:
        plist_file.write(plist_content)

    # LaunchAgent neu laden
    os.system(f"launchctl unload {launch_agent_path}")
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


def sanitize_filename(filename):
    """keine unzulässige Zeichen"""
    return re.sub(r'[\/:*?"<>|]', '_', filename)


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


def main():
    """Hauptverarbeitung"""
    try:
        process_scan_files(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")
        time.sleep(5)  # Warten
        main()  # Programm neu starten


if __name__ == "__main__":
    check_and_update_launch_agent()
    log_message("Programm gestartet.")
    while True:
        try:
            main()  # Hauptverarbeitung
        except Exception as e:
            log_message(f"Fehler im Hauptprogramm: {e}")
            time.sleep(5)  # Warten vor Neustart
