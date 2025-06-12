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
    # --- Colored Logging Setup ---
    class ColoredFormatter(logging.Formatter):
        RED = "\033[31m"
        YELLOW = "\033[33m"
        BLUE = "\033[34m"
        GREY = "\033[90m"
        GREEN = "\033[32m"
        RESET = "\033[0m"

        def format(self, record):
            if record.levelno == logging.ERROR:
                color = self.RED
            elif record.levelno == logging.WARNING:
                color = self.YELLOW
            elif record.levelno == logging.INFO:
                color = self.BLUE
            elif record.levelno == logging.DEBUG:
                color = self.GREY
            elif record.levelno == 25:  # SUCCESS
                color = self.GREEN
            else:
                color = self.RESET
            record.msg = f"{color}{record.msg}{self.RESET}"
            return super().format(record)

    SUCCESS_LEVEL = 25
    logging.addLevelName(SUCCESS_LEVEL, "SUCCESS")

    def success(self, message, *args, **kwargs):
        if logging.getLogger().isEnabledFor(SUCCESS_LEVEL):
            logging.getLogger()._log(SUCCESS_LEVEL, message, args, **kwargs)

    logging.Logger.success = success

    handler = logging.StreamHandler()
    handler.setFormatter(ColoredFormatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.getLogger().handlers = [handler]
    logging.getLogger().setLevel(logging.DEBUG)


configure_logging()


# Nouvelle fonction pour obtenir la date du jour avec gestion de la locale
def get_date_today(locale_str='fr_FR.UTF-8'):
    try:
        locale.setlocale(locale.LC_TIME, locale_str)
    except locale.Error as e:
        logging.warning(f"Impossible de définir la locale '{locale_str}' : {e}")
    return datetime.now().strftime('%d %B %Y')

# Nouvelle fonction pour initialiser les chemins de dossiers utilisés dans le script
def init_paths(base_dir):
    paths = {}
    paths['BASE_DIR'] = base_dir
    paths['template_dir'] = os.path.join(base_dir, 'utils', 'dossier_modele')
    paths['utils_dir'] = os.path.join(base_dir, 'utils')
    paths['input_dir'] = os.path.join(base_dir, 'documents')

    if not os.path.exists(paths['input_dir']):
        os.makedirs(paths['input_dir'])

    return paths

# Nouvelle fonction pour charger les données CSV
def load_csv_dataset(utils_dir):
    import glob
    csv_dir = os.path.join(utils_dir, 'csv')
    all_csv_paths = glob.glob(os.path.join(csv_dir, '*.csv'))
    csv_data = pd.concat([
        pd.read_csv(f, dtype=str, on_bad_lines='skip', keep_default_na=False)
        for f in all_csv_paths
    ], ignore_index=True)
    logging.info(f"{len(all_csv_paths)} fichiers CSV détectés :")
    for path in all_csv_paths:
        logging.info(f" - {os.path.basename(path)}")
    logging.info(f"Nombre total de lignes chargées : {len(csv_data)}")

    csv_data['Code postal'] = csv_data['Code postal'].str.zfill(5)
    csv_data['Images'] = csv_data['Images'].str.split('|').str[0]
    return csv_data

def load_department_mapping(utils_dir):
    departments_csv_path = os.path.join(utils_dir, 'departements-region.csv')
    departments_data = pd.read_csv(
        departments_csv_path,
        dtype={'num_dep': str}
    )
    departments_data['num_dep'] = departments_data['num_dep'].str.zfill(2)
    department_mapping = departments_data.set_index('num_dep')['dep_name'].to_dict()
    return department_mapping


# Nouvelle fonction pour générer les fiches individuelles
def generate_fiches(csv_data, paths, department_mapping, date_today):
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

    docx_files_by_folder = defaultdict(list)
    name_counter = {}

    logging.info("Début de la génération des fiches individuelles...")

    for index, row in csv_data.iterrows():
        try:
            numero_departement = row['Code postal'][:2]
            ville_upper = row['Ville'].upper()
            model_copy_dir = os.path.join(BASE_DIR, "dossiers_generes", f"{numero_departement} {ville_upper}")
            if not os.path.exists(model_copy_dir):
                shutil.copytree(template_dir, model_copy_dir)
            infractions_dir = os.path.join(model_copy_dir, '02 Infractions')
            photos_dir = os.path.join(model_copy_dir, '01 Photos')

            infraction_publicite = row.get('infraction_publicite')
            infraction_enseigne = row.get('infraction_enseigne')
            infraction_rlpi = row.get('infraction_rlpi')
            if infraction_publicite:
                infraction_text = infraction_publicite
                role_label = "Afficheur ou bénéficiaire"
            elif infraction_enseigne:
                infraction_text = infraction_enseigne
                role_label = "Annonceur"
                row['afficheur'] = row.get('annonceur', '')
            else:
                infraction_text = infraction_rlpi
                role_label = ""

            if row.get('afficheur_non_visible', '').strip().lower() == 'on':
                role_label = "non visible"

            categorie_text = row['Catégories (libellés)']
            preenseigne_text = ""
            match = re.search(r"(.*)« Les préenseignes sont soumises aux dispositions qui régissent la publicité » \(article L\.581-19\)", categorie_text)
            if match:
                preenseigne_text = "« Les préenseignes sont soumises aux dispositions qui régissent la publicité » (article L.581-19)"
                categorie_clean = match.group(1).strip()
            else:
                categorie_clean = categorie_text

            numero_rue_brut = str(row.get('Numéro', '')).strip()
            lat_int = row.get('Latitude', '').split('.')[0] if '.' in row.get('Latitude', '') else None
            rue_lower = row['Rue'].lower().strip()
            if (numero_rue_brut in ('', '-', 'nan') or
                    (lat_int and numero_rue_brut == lat_int) or
                    rue_lower.startswith('autoroute')):
                numero_rue = ''
            else:
                numero_rue = numero_rue_brut

            image_url = row['Images']
            image_path = os.path.join(photos_dir, f"image_{index}.jpg")
            default_image_path = os.path.join(photos_dir, 'default.jpg')

            if image_url:
                response = requests.get(image_url)
                if response.status_code == 200:
                    with open(image_path, 'wb') as f:
                        f.write(response.content)
                else:
                    logging.error(f"Échec du téléchargement de l'image: {image_url} avec le status code {response.status_code}")
                    image_path = default_image_path
            else:
                logging.error(f"Aucune URL d'image fournie pour la ligne {index}")
                image_path = default_image_path

            if not os.path.exists(default_image_path):
                default_img = Image.new('RGB', (500, 500), color='white')
                os.makedirs(os.path.dirname(default_image_path), exist_ok=True)
                default_img.save(default_image_path)

            department_name = department_mapping.get(numero_departement, 'Département Inconnu')
            afficheur, annonceur = (row['afficheur'].split(' - ') + [''])[:2]
            try:
                lat_short = f"{float(row['Latitude']):.5f}"
            except (ValueError, TypeError):
                lat_short = row['Latitude']
            try:
                lon_short = f"{float(row['Longitude']):.5f}"
            except (ValueError, TypeError):
                lon_short = row['Longitude']

            doc_template_path = os.path.join(utils_dir, 'fichev1.docx')
            doc = DocxTemplate(doc_template_path)
            context = {
                'Latitude': lat_short,
                'Longitude': lon_short,
                'gps': f"{lat_short}, {lon_short}",
                'my_image': InlineImage(doc, image_path, width=Mm(120)),
                'numero_de_rue': numero_rue,
                'rue': row['Rue'],
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
                'surface_estimee': f"surface estimée de {row['surface']} m²" if row['surface'].strip() else '',
            }

            try:
                doc.render(context)
            except Exception as e:
                logging.warning(f"Erreur de rendu avec données manquantes pour {row['Nom']} (index {index}): {e}")
                safe_context = {k: (v if isinstance(v, str) else str(v) if v is not None else '') for k, v in context.items()}
                doc.render(safe_context)

            infraction_paragraph = doc.add_paragraph()
            if infraction_text:
                process_html_content(infraction_paragraph, infraction_text)

            base_name = row['Nom']
            if base_name in name_counter:
                name_counter[base_name] += 1
                filename = f"{base_name}X{name_counter[base_name]:02d}.docx"
            else:
                name_counter[base_name] = 0
                filename = f"{base_name}.docx"

            modified_docx_path = os.path.join(infractions_dir, filename)
            doc.save(modified_docx_path)
            logging.info(f"Fichier DOCX créé: {modified_docx_path}")
            docx_files_by_folder[infractions_dir].append(modified_docx_path)
        except Exception as e:
            logging.error(f"Erreur lors de la génération de la fiche {row['Nom']} (index {index}): {e}")

    return docx_files_by_folder




from utils.courrier_infractions import generate_courrier

from utils.commune import fetch_commune_code
from utils.html_utils import process_html_content


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

docx_files_by_folder = generate_fiches(csv_data, paths, department_mapping, date_today)

# Génération des courriers pour chaque commune
def generate_courriers(csv_data, BASE_DIR, utils_dir):
    generated_communes = set((row['Code postal'][:2], row['Ville']) for _, row in csv_data.iterrows())

    for dep_code, ville in generated_communes:
        # Regrouper les lignes pour cette commune
        rows = [row for _, row in csv_data.iterrows() if (row['Code postal'][:2], row['Ville']) == (dep_code, ville)]
        courriers_dir = os.path.join(BASE_DIR, "dossiers_generes", f"{dep_code} {ville.upper()}", '03 Courriers')
        if not os.path.exists(courriers_dir):
            os.makedirs(courriers_dir)
        generate_courrier(rows, utils_dir, courriers_dir)

generate_courriers(csv_data, BASE_DIR, utils_dir)


# Nouvelle fonction pour fusionner les fichiers DOCX par commune
def merge_docx_per_commune(docx_files_by_folder):
    from utils.merge_docx import merge_docx_files
    for folder, files in docx_files_by_folder.items():
        if files:
            base_name = os.path.basename(files[0]).rsplit('-', 1)[0] + ".docx"
            combined_path = os.path.join(folder, base_name)
            merge_docx_files(files, combined_path)
            logging.info(f"Fichier combiné créé : {combined_path}")
            indiv_dir = os.path.join(folder, "indiv")
            os.makedirs(indiv_dir, exist_ok=True)
            for path in files:
                try:
                    shutil.move(path, os.path.join(indiv_dir, os.path.basename(path)))
                    logging.info(f"Fichier déplacé dans 'indiv' : {path}")
                except Exception as e:
                    logging.warning(f"Impossible de déplacer {path} : {e}")
from utils.merge_docx import merge_docx_files

merge_docx_per_commune(docx_files_by_folder)
