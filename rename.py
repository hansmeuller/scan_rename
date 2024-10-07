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


# Logdatei prüfen/create
def check_and_create_log_file():
    """create if not exist"""
    if not os.path.exists(log_file_path):
        with open(log_file_path, "w") as log_file:
            log_file.write(f"Logdatei erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")


# Einträge mit timestamp
def log_message(message):
    """message+timestamp"""
    check_and_create_log_file()
    clean_old_log_entries()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{timestamp} - {message}\n")


# lösche +=woche
def clean_old_log_entries():
    """Löscht einträge +=woche"""
    if os.path.exists(log_file_path):
        with open(log_file_path, "r") as log_file:
            lines = log_file.readlines()

        woche_alt = datetime.now() - timedelta(days=7)
        with open(log_file_path, "w") as log_file:
            for line in lines:
                try:
                    log_time = datetime.strptime(line.split(" - ")[0], "%Y-%m-%d %H:%M:%S")
                    if log_time > woche_alt:
                        log_file.write(line)
                except (ValueError, IndexError):
                    log_file.write(line)


# extrahiere pdfminer
def extract_text_with_format_from_pdf(pdf_path):
    """extrahiere"""
    formatted_text = []
    for page_layout in extract_pages(pdf_path):
        for element in page_layout:
            if isinstance(element, LTTextBoxHorizontal) or isinstance(element, LTTextLineHorizontal):
                for text_line in element:
                    text = text_line.get_text()
                    font_size = text_line.height  # schriftgröße
                    font_name = text_line.fontname if hasattr(text_line, 'fontname') else "unknown"

                    if font_name.lower().find("bold") != -1 or font_size > 12:  # fettdruck simulieren
                        formatted_text.append((text.strip(), "bold"))
                    else:
                        formatted_text.append((text.strip(), "normal"))

    return formatted_text


# suche fettgedrucktes
def get_subject_from_formatted_text(formatted_text):
    """fettgedruckt"""
    subject = None
    for line, format in formatted_text:
        if format == "bold":  # betreff ist fettgedruckt
            subject = line.strip()
            break
    return subject


# pfad
def process_scan_files(scans_folder):
    """pfad"""
    for file_name in os.listdir(scans_folder):
        if file_name.endswith(".pdf"):  # pdf only
            file_path = os.path.join(scans_folder, file_name)
            log_message(f"Verarbeite {file_name}...")

            # extrahiere
            formatted_text = extract_text_with_format_from_pdf(file_path)
            subject = get_subject_from_formatted_text(formatted_text)

            if subject:
                absender = file_name.split('_')[1]  # 2te stelle
                date = file_name.split('_')[0]  # datum first
                new_name = f"{absender}_{date}_{subject}"
                rename_file(file_path, absender, date, subject)
            else:
                log_message(f"Kein Betreff gefunden für {file_name}.")


# apple autostart
def check_and_update_launch_agent():
    """prüfe+aktualisieren"""
    launch_agent_path = os.path.expanduser("~/Library/LaunchAgents/com.meinprogramm.scanrenamer.plist")
    current_program_path = os.path.realpath(__file__)

    if os.path.exists(launch_agent_path):
        with open(launch_agent_path, "r") as plist_file:
            content = plist_file.read()

        if current_program_path in content:
            log_message(f"LaunchAgent ist aktuell. Programm wird von {current_program_path} ausgeführt.")
        else:
            log_message("LaunchAgent veraltet. Wird aktualisiert.")
            create_launch_agent(current_program_path)
    else:
        log_message("LaunchAgent existiert nicht. Erstellt neuen LaunchAgent.")
        create_launch_agent(current_program_path)


# LaunchAgent erstellen
def create_launch_agent(program_path):
    """Erstellt einen LaunchAgent für das Programm."""
    launch_agent_path = os.path.expanduser("~/Library/LaunchAgents/com.meinprogramm.scanrenamer.plist")

    plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
    <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
    <plist version="1.0">
    <dict>
        <key>Label</key>
        <string>com.meinprogramm.scanrenamer</string>

        <key>ProgramArguments</key>
        <array>
            <string>/usr/bin/python3</string>
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

    os.system(f"launchctl unload {launch_agent_path}")
    os.system(f"launchctl load {launch_agent_path}")
    log_message(f"LaunchAgent wurde erstellt und geladen für {program_path}.")


# check
def sanitize_filename(filename):
    """keine unzulässigen"""
    return re.sub(r'[\/:*?"<>|]', '_', filename)


# Datei umbenennen
def rename_file(old_path, absender, date, subject):
    """Absender_Datum_Betreff"""
    directory, old_file_name = os.path.split(old_path)
    file_extension = old_file_name.split('.')[-1]
    new_file_name = sanitize_filename(f"{absender}_{date}_{subject}.{file_extension}")
    new_file_path = os.path.join(directory, new_file_name)

    if not os.path.exists(new_file_path):
        os.rename(old_path, new_file_path)
        log_message(f"Datei umbenannt von {old_file_name} zu {new_file_name}")
    else:
        log_message(f"Fehler: Die Datei {new_file_path} existiert bereits.")


# Hauptverarbeitung
def main():
    """main"""
    try:
        process_scan_files(scans_folder)
    except Exception as e:
        log_message(f"Programm abgestürzt: {e}")
        time
