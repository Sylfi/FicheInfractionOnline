# main.py
import os
import locale
import logging
import re
from datetime import datetime
import requests
import shutil
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

def configure_logging():
    """
    Configure the logging system with colored output for different log levels.
    Defines a custom SUCCESS level and colors for ERROR, WARNING, INFO, DEBUG, and SUCCESS.
    """
    # --- Colored Logging Setup ---
    class ColoredFormatter(logging.Formatter):
        RED = "\033[31m"
        YELLOW = "\033[33m"
        BLUE = "\033[34m"
        GREY = "\033[90m"
        GREEN = "\033[32m"
        RESET = "\033[0m"

        def format(self, record):
            # Assign color based on the log level
            if record.levelno == logging.ERROR:
                color = self.RED
            elif record.levelno == logging.WARNING:
                color = self.YELLOW
            elif record.levelno == logging.INFO:
                color = self.BLUE
            elif record.levelno == logging.DEBUG:
                color = self.GREY
            elif record.levelno == 25:  # SUCCESS custom level
                color = self.GREEN
            else:
                color = self.RESET
            # Wrap the message with color codes
            record.msg = f"{color}{record.msg}{self.RESET}"
            return super().format(record)

    SUCCESS_LEVEL = 25
    logging.addLevelName(SUCCESS_LEVEL, "SUCCESS")

    def success(self, message, *args, **kwargs):
        # Custom method to log SUCCESS messages if enabled
        if logging.getLogger().isEnabledFor(SUCCESS_LEVEL):
            logging.getLogger()._log(SUCCESS_LEVEL, message, args, **kwargs)

    logging.Logger.success = success

    # Set up the stream handler with the colored formatter
    handler = logging.StreamHandler()
    handler.setFormatter(ColoredFormatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.getLogger().handlers = [handler]
    logging.getLogger().setLevel(logging.DEBUG)


configure_logging()


# Nouvelle fonction pour obtenir la date du jour avec gestion de la locale
def get_date_today(locale_str='fr_FR.UTF-8'):
    """
    Return the current date formatted as 'day month year' in French locale.
    Attempts to set the locale to French; logs a warning if unsuccessful.
    
    Parameters:
        locale_str (str): Locale string to set for date formatting.
        
    Returns:
        str: Formatted current date.
    """
    try:
        locale.setlocale(locale.LC_TIME, locale_str)
    except locale.Error as e:
        # Log warning if the specified locale is not available on the system
        logging.warning(f"Impossible de définir la locale '{locale_str}' : {e}")
    # Return formatted date string like '25 juin 2024'
    return datetime.now().strftime('%d %B %Y')

# Nouvelle fonction pour initialiser les chemins de dossiers utilisés dans le script
def init_paths(base_dir):
    """
    Initialize and return a dictionary of important directory paths used in the script.
    Ensures the input directory exists.
    
    Parameters:
        base_dir (str): Base directory where the script is located.
        
    Returns:
        dict: Paths for base, template, utils, and input directories.
    """
    paths = {}
    paths['BASE_DIR'] = base_dir
    paths['template_dir'] = os.path.join(base_dir, 'utils', 'dossier_modele')
    paths['utils_dir'] = os.path.join(base_dir, 'utils')
    paths['input_dir'] = os.path.join(base_dir, 'documents')

    # Create the input directory if it does not exist to avoid errors later
    if not os.path.exists(paths['input_dir']):
        os.makedirs(paths['input_dir'])

    return paths

# Nouvelle fonction pour charger les données CSV
def load_csv_dataset(utils_dir):
    """
    Load and concatenate all CSV files found in the 'csv' subdirectory of utils_dir.
    Cleans and processes certain columns for consistency.
    
    Parameters:
        utils_dir (str): Path to the utils directory containing CSV files.
        
    Returns:
        pd.DataFrame: Combined dataframe of all CSV data.
    """
    import glob
    csv_dir = os.path.join(utils_dir, 'csv')
    all_csv_paths = glob.glob(os.path.join(csv_dir, '*.csv'))
    # Read all CSV files as strings, skipping bad lines and avoiding NA interpretation
    csv_data = pd.concat([
        pd.read_csv(f, dtype=str, on_bad_lines='skip', keep_default_na=False)
        for f in all_csv_paths
    ], ignore_index=True)
    logging.info(f"{len(all_csv_paths)} fichiers CSV détectés :")
    for path in all_csv_paths:
        logging.info(f" - {os.path.basename(path)}")
    logging.info(f"Nombre total de lignes chargées : {len(csv_data)}")

    # Ensure postal codes have 5 digits, padding with zeros if needed
    csv_data['Code postal'] = csv_data['Code postal'].str.zfill(5)
    # For 'Images' column, keep only the first image URL (split by '|')
    csv_data['Images'] = csv_data['Images'].str.split('|').str[0]
    return csv_data

def load_department_mapping(utils_dir):
    """
    Load the department mapping from CSV file, mapping department numbers to names.
    Pads department numbers with zeros when necessary.
    
    Parameters:
        utils_dir (str): Path to utils directory containing 'departements-region.csv'.
        
    Returns:
        dict: Mapping of department number (str) to department name (str).
    """
    departments_csv_path = os.path.join(utils_dir, 'departements-region.csv')
    departments_data = pd.read_csv(
        departments_csv_path,
        dtype={'num_dep': str}
    )
    # Ensure department numbers are two digits, zero-padded
    departments_data['num_dep'] = departments_data['num_dep'].str.zfill(2)
    # Create a dictionary mapping 'num_dep' to 'dep_name'
    department_mapping = departments_data.set_index('num_dep')['dep_name'].to_dict()
    return department_mapping


# Nouvelle fonction pour générer les fiches individuelles
def generate_fiches(csv_data, paths, department_mapping, date_today):
    """
    Generate individual DOCX fiches (reports) for each row in the CSV data.
    Copies a template directory for each city/department and populates DOCX files with data.
    Downloads and embeds images, handles address cleaning, and manages naming conflicts.
    
    Parameters:
        csv_data (pd.DataFrame): DataFrame with all infraction data.
        paths (dict): Dictionary of directory paths.
        department_mapping (dict): Mapping from department codes to names.
        date_today (str): Current date formatted string.
        
    Returns:
        dict: Mapping from folder path to list of generated DOCX file paths.
    """
    from utils.html_utils import process_html_content
    from docxtpl import DocxTemplate, InlineImage
    from docx.shared import Mm
    from PIL import Image
    import re
    import requests
    import shutil
    from collections import defaultdict

    BASE_DIR = paths['BASE_DIR']
    template_dir = paths['template_dir']
    utils_dir = paths['utils_dir']

    # Dictionary to store generated DOCX files by their folder
    docx_files_by_folder = defaultdict(list)
    # Counter to track duplicate filenames for naming conflicts
    name_counter = {}

    logging.info("Début de la génération des fiches individuelles...")

    # Iterate over each row to create individual fiches
    for index, row in csv_data.iterrows():
        try:
            # Extract department code from postal code (first two digits)
            numero_departement = row['Code postal'][:2]
            # Use uppercase city name for folder naming
            ville_upper = row['Ville'].upper()
            # Define path to copy the template folder for this city/department
            model_copy_dir = os.path.join(BASE_DIR, "dossiers_generes", f"{numero_departement} {ville_upper}")
            # Copy template directory if it doesn't exist yet
            if not os.path.exists(model_copy_dir):
                shutil.copytree(template_dir, model_copy_dir)
            infractions_dir = os.path.join(model_copy_dir, '02 Infractions')
            photos_dir = os.path.join(model_copy_dir, '01 Photos')

            # Determine which infraction text and role label to use based on available data
            infraction_publicite = row.get('infraction_publicite')
            infraction_enseigne = row.get('infraction_enseigne')
            infraction_rlpi = row.get('infraction_rlpi')
            if infraction_publicite:
                infraction_text = infraction_publicite
                role_label = "Afficheur ou bénéficiaire"
            elif infraction_enseigne:
                infraction_text = infraction_enseigne
                role_label = "Annonceur"
                # If infraction_enseigne is used, assign afficheur from annonceur field if available
                row['afficheur'] = row.get('annonceur', '')
            else:
                infraction_text = infraction_rlpi
                role_label = ""

            # If 'afficheur_non_visible' flag is set to 'on', override role label
            if row.get('afficheur_non_visible', '').strip().lower() == 'on':
                role_label = "non visible"

            # Process category text to extract and clean preenseigne clause if present
            categorie_text = row['Catégories (libellés)']
            preenseigne_text = ""
            # Search for specific phrase indicating preenseigne legal clause
            match = re.search(r"(.*)« Les préenseignes sont soumises aux dispositions qui régissent la publicité » \(article L\.581-19\)", categorie_text)
            if match:
                # Extract and save the legal clause separately
                preenseigne_text = "« Les préenseignes sont soumises aux dispositions qui régissent la publicité » (article L.581-19)"
                # Clean category text by removing the clause
                categorie_clean = match.group(1).strip()
            else:
                categorie_clean = categorie_text

            # Clean and validate street number:
            numero_rue_brut = str(row.get('Numéro', '')).strip()
            lat_int = row.get('Latitude', '').split('.')[0] if '.' in row.get('Latitude', '') else None
            rue_lower = row['Rue'].lower().strip()
            # Conditions to discard the number:
            # - Empty, dash, or 'nan' string
            # - Number matches integer part of latitude (likely erroneous)
            # - Street name starts with 'autoroute' (highway, no street number)
            if (numero_rue_brut in ('', '-', 'nan') or
                    (lat_int and numero_rue_brut == lat_int) or
                    rue_lower.startswith('autoroute')):
                numero_rue = ''
            else:
                numero_rue = numero_rue_brut

            # Handle image downloading and fallback to default image if needed
            image_url = row['Images']
            image_path = os.path.join(photos_dir, f"image_{index}.jpg")
            default_image_path = os.path.join(photos_dir, 'default.jpg')

            if image_url:
                response = requests.get(image_url)
                if response.status_code == 200:
                    # Save the downloaded image locally
                    with open(image_path, 'wb') as f:
                        f.write(response.content)
                else:
                    # Log error and use default image if download fails
                    logging.error(f"Échec du téléchargement de l'image: {image_url} avec le status code {response.status_code}")
                    image_path = default_image_path
            else:
                # Log error and use default image if no image URL provided
                logging.error(f"Aucune URL d'image fournie pour la ligne {index}")
                image_path = default_image_path

            # If default image does not exist, create a blank white image as placeholder
            if not os.path.exists(default_image_path):
                default_img = Image.new('RGB', (500, 500), color='white')
                os.makedirs(os.path.dirname(default_image_path), exist_ok=True)
                default_img.save(default_image_path)

            # Get department name from mapping, fallback if unknown
            department_name = department_mapping.get(numero_departement, 'Département Inconnu')
            # Split 'afficheur' field into afficheur and annonceur if separated by ' - '
            afficheur, annonceur = (row['afficheur'].split(' - ') + [''])[:2]
            # Format latitude and longitude to 5 decimal places for display
            try:
                lat_short = f"{float(row['Latitude']):.5f}"
            except (ValueError, TypeError):
                lat_short = row['Latitude']
            try:
                lon_short = f"{float(row['Longitude']):.5f}"
            except (ValueError, TypeError):
                lon_short = row['Longitude']

            # Load the DOCX template for the fiche
            doc_template_path = os.path.join(utils_dir, 'fichev1.docx')
            doc = DocxTemplate(doc_template_path)
            # Prepare the context dictionary for template rendering
            context = {
                'Latitude': lat_short,
                'Longitude': lon_short,
                'gps': f"{lat_short}, {lon_short}",
                'my_image': InlineImage(doc, image_path, width=Mm(120)),
                'numero_de_rue': numero_rue,
                'rue': row['Rue'],
                # Location is street number + street name if number present, else just street name
                'localisation': f"{numero_rue} {row['Rue']}".strip() if numero_rue else row['Rue'],
                'numero_de_fiche': row['Nom'],
                'code_fiche': row['Nom'],
                'nom_commune': f"{row['Ville']}",
                'code_postal': row['Code postal'],
                'department_name': department_name,
                'numero_departement': numero_departement,
                'date_today': date_today,
                'afficheur': afficheur.strip(),
                'annonceur': annonceur.strip(),
                'type_dispositif': categorie_clean,
                'type_infraction': '',
                'préenseignes': preenseigne_text,
                'annonceurafficheur': role_label,
                # Include surface estimate text only if surface is non-empty after stripping whitespace
                'surface_estimee': f"surface estimée de {row['surface']} m²" if row['surface'].strip() else '',
            }

            try:
                # Attempt to render the template with the context
                doc.render(context)
            except Exception as e:
                # If rendering fails (likely due to missing data), try converting all values to strings
                logging.warning(f"Erreur de rendu avec données manquantes pour {row['Nom']} (index {index}): {e}")
                safe_context = {k: (v if isinstance(v, str) else str(v) if v is not None else '') for k, v in context.items()}
                doc.render(safe_context)

            # Add the infraction text as formatted HTML content in a new paragraph
            infraction_paragraph = doc.add_paragraph()
            if infraction_text:
                process_html_content(infraction_paragraph, infraction_text)

            # Handle filename conflicts: if base name already used, append 'X' and a counter
            base_name = row['Nom']
            if base_name in name_counter:
                name_counter[base_name] += 1
                filename = f"{base_name}X{name_counter[base_name]:02d}.docx"
            else:
                name_counter[base_name] = 0
                filename = f"{base_name}.docx"

            # Save the rendered DOCX file in the infractions directory
            modified_docx_path = os.path.join(infractions_dir, filename)
            doc.save(modified_docx_path)
            logging.info(f"Fichier DOCX créé: {modified_docx_path}")
            # Track the saved file by folder for later merging
            docx_files_by_folder[infractions_dir].append(modified_docx_path)
        except Exception as e:
            # Log any errors during fiche generation without stopping the loop
            logging.error(f"Erreur lors de la génération de la fiche {row['Nom']} (index {index}): {e}")

    return docx_files_by_folder




from utils.courrier_infractions import generate_courrier

from utils.commune import fetch_commune_code
from utils.html_utils import process_html_content



# Génération des courriers pour chaque commune
def generate_courriers(csv_data, BASE_DIR, utils_dir):
    """
    Generate letters ('courriers') for each unique commune (department code + city) found in the CSV data.
    Groups rows by commune and calls the courrier generation utility.
    
    Parameters:
        csv_data (pd.DataFrame): DataFrame with all infraction data.
        BASE_DIR (str): Base directory for generated dossiers.
        utils_dir (str): Directory containing utility scripts.
    """
    # Create a set of unique (department code, city) tuples
    generated_communes = set((row['Code postal'][:2], row['Ville']) for _, row in csv_data.iterrows())

    for dep_code, ville in generated_communes:
        # Filter rows belonging to the current commune
        rows = [row for _, row in csv_data.iterrows() if (row['Code postal'][:2], row['Ville']) == (dep_code, ville)]
        courriers_dir = os.path.join(BASE_DIR, "dossiers_generes", f"{dep_code} {ville.upper()}", '03 Courriers')
        # Create the courriers directory if it doesn't exist
        if not os.path.exists(courriers_dir):
            os.makedirs(courriers_dir)
        # Generate the courrier documents for this commune
        generate_courrier(rows, utils_dir, courriers_dir)

# Nouvelle fonction pour fusionner les fichiers DOCX par commune
def merge_docx_per_commune(docx_files_by_folder):
    """
    Merge individual DOCX files into a single combined DOCX file per folder (commune).
    Moves individual files into an 'indiv' subfolder after merging.
    
    Parameters:
        docx_files_by_folder (dict): Mapping from folder path to list of DOCX file paths.
    """
    from utils.merge_docx import merge_docx_files
    for folder, files in docx_files_by_folder.items():
        if files:
            # Determine base name for combined file by stripping suffix after last hyphen
            base_name = os.path.basename(files[0]).rsplit('-', 1)[0] + ".docx"
            combined_path = os.path.join(folder, base_name)
            # Merge the DOCX files into one combined document
            merge_docx_files(files, combined_path)
            logging.info(f"Fichier combiné créé : {combined_path}")
            # Create an 'indiv' subfolder to store individual files post-merge
            indiv_dir = os.path.join(folder, "indiv")
            os.makedirs(indiv_dir, exist_ok=True)
            for path in files:
                try:
                    # Move each individual file into the 'indiv' folder
                    shutil.move(path, os.path.join(indiv_dir, os.path.basename(path)))
                    logging.info(f"Fichier déplacé dans 'indiv' : {path}")
                except Exception as e:
                    # Log a warning if moving fails but continue processing
                    logging.warning(f"Impossible de déplacer {path} : {e}")
from utils.merge_docx import merge_docx_files


# ========================== MAIN WRAPPER ==========================
def main():
    """
    Main function orchestrating the workflow:
    1. Sets locale and gets current date.
    2. Initializes directory paths.
    3. Loads CSV datasets and department mappings.
    4. Generates individual fiches (reports).
    5. Generates courriers (letters) for each commune.
    6. Merges individual DOCX files per commune.
    """
    # Définir la locale et obtenir la date du jour
    date_today = get_date_today()

    # Initialisation des chemins de dossiers
    paths = init_paths(os.path.dirname(os.path.abspath(__file__)))
    BASE_DIR = paths['BASE_DIR']
    template_dir = paths['template_dir']
    utils_dir = paths['utils_dir']
    input_dir = paths['input_dir']

    # Charger les données CSV
    csv_data = load_csv_dataset(utils_dir)

    # Charger le mapping des départements
    department_mapping = load_department_mapping(utils_dir)

    # Générer les fiches
    docx_files_by_folder = generate_fiches(csv_data, paths, department_mapping, date_today)

    # Générer les courriers
    generate_courriers(csv_data, BASE_DIR, utils_dir)

    # Fusionner les fichiers DOCX
    merge_docx_per_commune(docx_files_by_folder)

if __name__ == "__main__":
    main()
