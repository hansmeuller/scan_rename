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
