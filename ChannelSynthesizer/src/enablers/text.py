import os
import json
import re

import fitz
from parsers.providers.orange import parse_orange_pdf
from parsers.providers.voo import parse_voo_pdf
from parsers.providers.telenet import parse_telenet_pdf
from parsers.all_sections_parser import detect_provider_and_year
from utils import read_section_names

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '../'))
CONFIG_PATH = os.path.join(BASE_DIR, '.config/page_selection.json')

def load_page_selection() -> dict:
    """
    cette fonction charge la configuration de la sélection de pages à partir d'un fichier json.
    elle retourne un dictionnaire de la selection de pages
    """
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def get_pages_to_process(pdf_path: str, total_pages: int) -> list:
    """
    cette fonction obtient les pages à traiter selon la configuration
    pdf_path c'est le chemin du fichier pdf et total_pages c'est le nombre total de pages dans le pdf
    elle retourne une liste de numéros de pages à traiter
    """
    page_selection = load_page_selection()
    filename = os.path.basename(pdf_path)
    return page_selection.get(filename, list(range(1, total_pages + 1)))

def add_tv_radio_codes(tsv_path, section_names):
    """
    cette fonction ajoute les codes tv/radio dans le fichier tsv
    section_names sont les noms de section chargées pour vérifier la catégorie
    """
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    current_section = None
    processed_lines = []

    for line in lines:
        stripped_line = line.strip()

        # mise à jour de la section actuelle si la ligne est un nom de section
        if stripped_line in section_names:
            current_section = stripped_line

        # sauter les lignes qui sont seulement des chiffres ou des noms de section
        if stripped_line.isdigit() or stripped_line in section_names:
            processed_lines.append(line)
            continue

        # verifier si la section actuelle devrait etre catégorisée comme radio
        is_radio_section = (
            current_section and re.search(r'radio|stingray|music|radios|radiozenders|CHAÎNES DE MUSIQUE', current_section, re.IGNORECASE)
        )

        # corriger la categorisation si la ligne a deja un code tv/r incorrect
        if re.search(r'\s(R|TV)$', stripped_line):
            stripped_line = re.sub(r'\s(R|TV)$', '', stripped_line)

        # ajouter le code correct
        if is_radio_section:
            processed_lines.append(f"{stripped_line} R\n")
        else:
            processed_lines.append(f"{stripped_line} TV\n")

    # gérer des cas spécifiques ou des chaines devraient être tv au lieu de radio
    channels_to_correct = [
        "National Geographic", "Ketnet", "STAR channel", "Plattelands TV", "vtm Gold",
        "BBC Entertainment", "Disney Channel VL", "BBC First", "Nickelodeon NL", "Nick Jr NL",
        "Nickelodeon Ukraine", "Disney JR NL", "Play6", "MENT TV", "Q-music", "Play Crime",
        "MTV", "TLC", "Comedy Central", "Eclips TV", "VTM non stop dokters", "History",
        "Play 7", "Cartoon Network", "Vlaams Parlement TV", "ID", "OUTtv", "Play Sports Info",
        "Al Aoula Europe", "2M Monde", "Al Maghreb TV", "TRT Turk", "MBC", "TV Polonia",
        "Rai Uno", "Rai Due", "Rai Tre", "Mediaset Italia", "TVE Internacional",
        "The Israëli Network", "BBC One", "BBC Two", "NPO 1", "NPO 2", "NPO 3", "ARD", "ZDF", "VOX"
    ]

    final_lines = []
    for line in processed_lines:
        # vérifier si la ligne correspond à une des chaines à corriger en tv
        if any(channel in line for channel in channels_to_correct):
            final_lines.append(re.sub(r' R$', ' TV', line))
        else:
            final_lines.append(line)

    with open(tsv_path, 'w', encoding='utf-8') as f:
        f.writelines(final_lines)

    print(f"codes tv/radio traitées et enregistrées dans {tsv_path}")

def process_pdfs(directory):
    """
    cette fonction traite tous les fichiers pdf dans le répertoire donné.
    directory c'est le répertoire contenant les fichiers pdf
    """
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            try:
                provider, year = detect_provider_and_year(pdf_path)
                document = fitz.open(pdf_path)
                total_pages = document.page_count
                document.close()


                section_file = os.path.join(BASE_DIR, 'outputs/section', os.path.splitext(filename)[0] + '_sections.tsv')
                tsv_path = os.path.join(BASE_DIR, 'outputs/text', os.path.splitext(filename)[0] + '_text.tsv')

                # charger les noms de section si disponible
                section_names = []

                if os.path.exists(section_file):
                    section_names = read_section_names(section_file)

                # parser le pdf basé sur le provider
                if provider == "VOO":
                    parse_voo_pdf(pdf_path)
                elif provider == "Telenet":
                    pages_to_process = get_pages_to_process(pdf_path, total_pages)  # passer total_pages ici
                    parse_telenet_pdf(pdf_path, pages_to_process)
                elif provider == "Orange":
                    parse_orange_pdf(pdf_path, section_names)  # passer section_names ici
                else:
                    print(f"provider non supporté {provider} pour le fichier {filename}")

                # appliquer le marquage tv/radio au fichier tsv
                add_tv_radio_codes(tsv_path, section_names)

            except ValueError as e:
                print(f"erreur en traitant {filename}: {e}")

if __name__ == "__main__":
    input_directory = os.path.join(BASE_DIR, 'inputs/pdf')
    process_pdfs(input_directory)
