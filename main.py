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
from typing import Optional


from utils.courrier_infractions import generate_courrier

from utils.commune import fetch_commune_code
from utils.html_utils import process_html_content

# Configuration de logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

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
csv_file = os.path.join(utils_dir, 'document.csv')
csv_data = pd.read_csv(
    csv_file,
    dtype={'Code postal': str},   # Conserver les zéros en tête
    on_bad_lines='skip'
)
csv_data = csv_data.fillna('').infer_objects(copy=False)
csv_data['Code postal'] = csv_data['Code postal'].str.zfill(5)
csv_data = csv_data.applymap(str)

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

docx_files = []
name_counter = {}
current_dir = None

# Traitement des données et génération des documents
for index, row in csv_data.iterrows():
    # Préparer le dossier de sortie pour cette commune
    numero_departement = row['Code postal'][:2]
    ville_upper = row['Ville'].upper()
    model_copy_dir = os.path.join(BASE_DIR, "dossiers_generes", f"{numero_departement} {ville_upper}")
    # Ne copier le modèle qu'une seule fois par commune
    if current_dir != model_copy_dir:
        if os.path.exists(model_copy_dir):
            shutil.rmtree(model_copy_dir)
        shutil.copytree(template_dir, model_copy_dir)
        current_dir = model_copy_dir
    infractions_dir = os.path.join(model_copy_dir, '02 Infractions')
    photos_dir = os.path.join(model_copy_dir, '01 Photos')
    # Détermination de l'infraction et du rôle associé selon la colonne utilisée
    infraction_publicite = row.get('infraction_publicite')
    infraction_enseigne = row.get('infraction_enseigne')
    infraction_rlpi = row.get('infraction_rlpi')
    if infraction_publicite:
        infraction_text = infraction_publicite
        role_label = "Afficheur"
    elif infraction_enseigne:
        infraction_text = infraction_enseigne
        role_label = "Annonceur"
    else:
        infraction_text = infraction_rlpi
        role_label = ""

    if row.get('afficheur_non_visible_1', '').strip().lower() == 'on':
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
    try:
        lat_int = str(int(float(row.get('Latitude', 0))))
    except Exception:
        lat_int = None

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
        default_img.save(default_image_path)

    numero_departement = row['Code postal'][:2]
    department_name = department_mapping.get(numero_departement, 'Département Inconnu')
    afficheur, annonceur = (row['afficheur'].split(' - ') + [''])[:2]

    doc_template_path = os.path.join(utils_dir, 'fichev1.docx')
    doc = DocxTemplate(doc_template_path)
    context = {
        'my_image': InlineImage(doc, image_path, width=Mm(120)),
        'numero_de_rue': numero_rue,
        'rue': row['Rue'],
        'localisation': f"{numero_rue} {row['Rue']}".strip() if numero_rue else row['Rue'],
        'numero_de_fiche': row['Nom'],
        'code_fiche': row['Nom'],
        'nom_commune': row['Ville'],
        'code_postal': row['Code postal'],
        'department_name': department_name,
        'numero_departement': numero_departement,
        'date_today': date_today,
        'gps': f"{float(row['Latitude']):.5f}, {float(row['Longitude']):.5f}",
        'afficheur': afficheur.strip(),
        'annonceur': annonceur.strip(),
        'type_dispositif': categorie_clean,
        'type_infraction': '',
        'préenseignes': preenseigne_text,
        'annonceurafficheur': role_label,
    }

    doc.render(context)
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
    docx_files.append(modified_docx_path)

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
