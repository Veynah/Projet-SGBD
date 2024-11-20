import os


def collect_code_to_file(output_file, extensions=None, exclude_dirs=None):
    """
    Parcourt tous les fichiers dans le répertoire courant et écrit leur contenu dans un fichier de sortie.

    :param output_file: Nom du fichier de sortie.
    :param extensions: Liste d'extensions de fichiers à inclure (par exemple ['.py', '.js']).
    :param exclude_dirs: Liste de répertoires à exclure (par exemple ['venv', '__pycache__']).
    """
    # Extensions par défaut si non spécifiées
    extensions = extensions or [".py"]
    exclude_dirs = set(exclude_dirs or [])

    current_directory = os.getcwd()  # Récupère le répertoire courant

    with open(output_file, "w", encoding="utf-8") as outfile:
        for root, dirs, files in os.walk(current_directory):
            # Filtrer les répertoires exclus
            dirs[:] = [d for d in dirs if d not in exclude_dirs]

            for file in files:
                if any(file.endswith(ext) for ext in extensions):
                    file_path = os.path.join(root, file)
                    try:
                        with open(file_path, "r", encoding="utf-8") as infile:
                            outfile.write(f"# File: {file_path}\n")
                            outfile.write(infile.read())
                            outfile.write("\n\n")  # Séparation entre les fichiers
                    except Exception as e:
                        print(f"Impossible de lire {file_path}: {e}")

    print(f"Tout le code a été collecté dans {output_file}")


if __name__ == "__main__":
    output_file = "code_complet.txt"
    extensions = [".py", ".js", ".html", ".css"]  # Extensions de fichiers à inclure
    exclude_dirs = ["venv", "__pycache__", ".git"]  # Répertoires à exclure

    collect_code_to_file(output_file, extensions, exclude_dirs)
