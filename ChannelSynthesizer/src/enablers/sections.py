import os

from parsers.all_sections_parser import extract_text, parse, get_provider_colors, detect_provider_and_year, get_pages_to_process, remove_redundant_sections, save_sections


BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '../'))

def process(folder_path: str) -> None:
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            path = os.path.join(folder_path, file)
            try:
                provider, year = detect_provider_and_year(path)
                colors = get_provider_colors(provider)
                pages = get_pages_to_process(path)

                all_sections = []

                for page_number in pages:
                    text, max_size = extract_text(path, colors, provider, page_number)
                    sections = parse(text, provider, max_size)
                    all_sections.extend(sections)

                all_sections = remove_redundant_sections(all_sections)

                output_path = os.path.join(BASE_DIR, 'outputs/section')
                save_sections(path, all_sections, output_dir=output_path)
                print(f"Saved sections for {file} for provider {provider} and year {year}")

            except ValueError as e:
                print(f"Error processing {file}: {e}")

if __name__ == "__main__":
    folder = os.path.join(BASE_DIR, 'inputs/pdf')
    process(folder)
