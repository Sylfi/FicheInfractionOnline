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




from utils.courrier_infractions import generate_courrier

from utils.commune import fetch_commune_code
from utils.html_utils import process_html_content


# Définir la locale
locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

# # Demander si on utilise la date du jour
# user_choice = input("Voulez-vous utiliser la date du jour ? (O/n): ")
# if user_choice.lower() in ['o', 'oui', '']:
#     date_today = datetime.now().strftime('%d %B %Y')
# else:
#     date_str = input("Entrez la date au format JJ/MM/AA: ")
#     try:
#         dt = datetime.strptime(date_str, "%d/%m/%y")
#         date_today = dt.strftime('%d %B %Y')
#     except Exception as e:
#         logging.error("Format incorrect, utilisation de la date du jour.")
#         date_today = datetime.now().strftime('%d %B %Y')
date_today = datetime.now().strftime('%d %B %Y')  # Toujours la date du jour

# Chemin vers le dossier 'documents' (relatif au script)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Dossier modèle à copier pour chaque commune
template_dir = os.path.join(BASE_DIR, 'utils', 'dossier_modele')
utils_dir = os.path.join(BASE_DIR, 'utils')
input_dir = os.path.join(BASE_DIR, 'documents')

if not os.path.exists(input_dir):
    os.makedirs(input_dir)

 # Charger les données CSV
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

# Diviser la chaîne de caractères pour les images
csv_data['Images'] = csv_data['Images'].str.split('|').str[0]


# Charger le mapping des départements
departments_csv_path = os.path.join(utils_dir, 'departements-region.csv')
# On lit num_dep comme chaîne et on garantit 2 caractères (ex. "01", "02", ...)
departments_data = pd.read_csv(
    departments_csv_path,
    dtype={'num_dep': str}    # Conserver les zéros en tête
)
departments_data['num_dep'] = departments_data['num_dep'].str.zfill(2)
department_mapping = departments_data.set_index('num_dep')['dep_name'].to_dict()

from collections import defaultdict
docx_files_by_folder = defaultdict(list)
name_counter = {}

logging.info("Début de la génération des fiches individuelles...")

# Traitement des données et génération des documents
for index, row in csv_data.iterrows():
    try:
        # Préparer le dossier de sortie pour cette commune
        numero_departement = row['Code postal'][:2]
        ville_upper = row['Ville'].upper()
        model_copy_dir = os.path.join(BASE_DIR, "dossiers_generes", f"{numero_departement} {ville_upper}")
        # Ne copier le modèle qu'une seule fois par commune
        if not os.path.exists(model_copy_dir):
            shutil.copytree(template_dir, model_copy_dir)
        infractions_dir = os.path.join(model_copy_dir, '02 Infractions')
        photos_dir = os.path.join(model_copy_dir, '01 Photos')
        # Détermination de l'infraction et du rôle associé selon la colonne utilisée
        infraction_publicite = row.get('infraction_publicite')
        infraction_enseigne = row.get('infraction_enseigne')
        infraction_rlpi = row.get('infraction_rlpi')
        if infraction_publicite:
            infraction_text = infraction_publicite
            role_label = "Afficheur ou bénéficiaire"
        elif infraction_enseigne:
            infraction_text = infraction_enseigne
            role_label = "Annonceur"
            # Si c'est une enseigne, on considère que c'est l'annonceur qui est affiché
            row['afficheur'] = row.get('annonceur', '')
        else:
            infraction_text = infraction_rlpi
            role_label = ""

        if row.get('afficheur_non_visible', '').strip().lower() == 'on':
            role_label = "non visible"

        # Traitement de la catégorie
        categorie_text = row['Catégories (libellés)']
        preenseigne_text = ""
        match = re.search(r"(.*)« Les préenseignes sont soumises aux dispositions qui régissent la publicité » \(article L\.581-19\)", categorie_text)
        if match:
            preenseigne_text = "« Les préenseignes sont soumises aux dispositions qui régissent la publicité » (article L.581-19)"
            categorie_clean = match.group(1).strip()
        else:
            categorie_clean = categorie_text

        # --- Nettoyage du numéro de rue ---
        numero_rue_brut = str(row.get('Numéro', '')).strip()

        # Partie entière de la latitude (ex. 44 pour 44,14114)
        lat_int = row.get('Latitude', '').split('.')[0] if '.' in row.get('Latitude', '') else None

        rue_lower = row['Rue'].lower().strip()

        # On ignore le numéro si :
        # 1) champ vide / '-' / 'nan'
        # 2) il correspond à la partie entière de la latitude (bug Gogocarto)
        # 3) la voie commence par « Autoroute »  (on ne met jamais de numéro)
        if (numero_rue_brut in ('', '-', 'nan') or
                (lat_int and numero_rue_brut == lat_int) or
                rue_lower.startswith('autoroute')):
            numero_rue = ''
        else:
            numero_rue = numero_rue_brut

        # Téléchargement de l'image
        image_url = row['Images']
        image_path = os.path.join(photos_dir, f"image_{index}.jpg")
        default_image_path = os.path.join(photos_dir, 'default.jpg')  # Image par défaut

        if image_url:
            response = requests.get(image_url)
            if response.status_code == 200:
                with open(image_path, 'wb') as f:
                    f.write(response.content)
            else:
                logging.error(f"Échec du téléchargement de l'image: {image_url} avec le status code {response.status_code}")
                image_path = default_image_path  # Utiliser une image par défaut
        else:
            logging.error(f"Aucune URL d'image fournie pour la ligne {index}")
            image_path = default_image_path  # Utiliser une image par défaut

        # Vérifier si l'image par défaut existe, sinon, créer un carré blanc
        if not os.path.exists(default_image_path):
            from PIL import Image
            default_img = Image.new('RGB', (500, 500), color='white')
            os.makedirs(os.path.dirname(default_image_path), exist_ok=True)
            default_img.save(default_image_path)

        numero_departement = row['Code postal'][:2]
        department_name = department_mapping.get(numero_departement, 'Département Inconnu')
        afficheur, annonceur = (row['afficheur'].split(' - ') + [''])[:2]

        # Truncate Latitude and Longitude to 5 decimal places for context
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
            # Préfixe de test pour vérifier la prise en compte des modifications
            'nom_commune': f"X{row['Ville']}",
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
        }

        try:
            doc.render(context)
        except Exception as e:
            logging.warning(f"Erreur de rendu avec données manquantes pour {row['Nom']} (index {index}): {e}")
            # Remplacer les valeurs non-string ou None par chaînes vides
            safe_context = {k: (v if isinstance(v, str) else str(v) if v is not None else '') for k, v in context.items()}
            doc.render(safe_context)
        # Ajout du texte des infractions avec le formatage correct et la police Arial
        infraction_paragraph = doc.add_paragraph()
        if infraction_text:
            process_html_content(infraction_paragraph, infraction_text)

        # Gérer les noms de fichiers en cas de doublons
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

logging.info("Traitement terminé.")

# Génération des courriers pour chaque commune
generated_communes = set((row['Code postal'][:2], row['Ville']) for _, row in csv_data.iterrows())

for dep_code, ville in generated_communes:
    # Regrouper les lignes pour cette commune
    rows = [row for _, row in csv_data.iterrows() if (row['Code postal'][:2], row['Ville']) == (dep_code, ville)]
    courriers_dir = os.path.join(BASE_DIR, "dossiers_generes", f"{dep_code} {ville.upper()}", '03 Courriers')
    if not os.path.exists(courriers_dir):
        os.makedirs(courriers_dir)
    generate_courrier(rows, utils_dir, courriers_dir)

from utils.merge_docx import merge_docx_files

for folder, files in docx_files_by_folder.items():
    if files:
        base_name = os.path.basename(files[0]).rsplit('-', 1)[0] + ".docx"
        combined_path = os.path.join(folder, base_name)
        merge_docx_files(files, combined_path)
        logging.info(f"Fichier combiné créé : {combined_path}")
        # Suppression des fichiers individuels après fusion réussie
        for path in files:
            try:
                os.remove(path)
                logging.info(f"Fichier supprimé : {path}")
            except Exception as e:
                logging.warning(f"Impossible de supprimer {path} : {e}")
