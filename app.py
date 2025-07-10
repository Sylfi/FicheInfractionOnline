# app.py
from flask import Flask, request, render_template_string, send_file
import os
import zipfile
import shutil
import tempfile
import pandas as pd

# On importe tes fonctions déjà présentes dans main.py
from main import init_paths, load_department_mapping, generate_fiches, generate_courriers, merge_docx_per_commune, get_date_today

# --------------------------------------------------------
# 1. Initialiser l'application Flask
# --------------------------------------------------------
app = Flask(__name__)

# --------------------------------------------------------
# 2. Définir la page d'accueil (route "/")
#    qui affiche un formulaire HTML simple
# --------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    return render_template_string("""
    <h2>Générateur de fiches infractions</h2>
    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="csvfile" accept=".csv" required><br><br>
        <button type="submit">Lancer le traitement</button>
    </form>
    """)

# --------------------------------------------------------
# 3. Définir la route "/process" qui gère le traitement
# --------------------------------------------------------
@app.route("/process", methods=["POST"])
def process():
    # -----------------------------------------------------
    # a. Créer un dossier temporaire unique pour cette session
    #    (chaque utilisateur ou usage aura son propre dossier isolé)
    # -----------------------------------------------------
    temp_dir = tempfile.mkdtemp()

    # -----------------------------------------------------
    # b. Enregistrer le fichier CSV reçu depuis le formulaire
    # -----------------------------------------------------
    csv_file = request.files['csvfile']
    csv_path = os.path.join(temp_dir, 'data.csv')
    csv_file.save(csv_path)

    # -----------------------------------------------------
    # c. Initialiser tes chemins de travail en utilisant ton init_paths
    #    Cela créera un import_csv dans temp_dir etc.
    # -----------------------------------------------------
    # Copier le dossier utils local vers le dossier temporaire
    shutil.copytree('utils', os.path.join(temp_dir, 'utils'))
    paths = init_paths(temp_dir)

    # -----------------------------------------------------
    # d. Déplacer le CSV vers import_csv comme tu le fais d'habitude
    # -----------------------------------------------------
    shutil.move(csv_path, os.path.join(paths['import_csv_dir'], 'data.csv'))

    # -----------------------------------------------------
    # e. Charger les données et ton mapping
    # -----------------------------------------------------
    csv_data = pd.read_csv(
        os.path.join(paths['import_csv_dir'], 'data.csv'),
        dtype=str, keep_default_na=False
    )
    department_mapping = load_department_mapping(paths['utils_dir'])
    date_today = get_date_today()

    # -----------------------------------------------------
    # f. Appeler exactement ton pipeline habituel
    # -----------------------------------------------------
    docx_files_by_folder = generate_fiches(csv_data, paths, department_mapping, date_today)
    generate_courriers(csv_data, paths['BASE_DIR'], paths['utils_dir'])
    merge_docx_per_commune(docx_files_by_folder)

    # -----------------------------------------------------
    # g. Créer un ZIP avec le dossier 'dossiers_generes'
    # -----------------------------------------------------
    zip_path = os.path.join(temp_dir, "resultats.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(os.path.join(temp_dir, "dossiers_generes")):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, os.path.join(temp_dir, "dossiers_generes"))
                zipf.write(file_path, arcname=arcname)

    # -----------------------------------------------------
    # h. Retourner le ZIP à télécharger directement dans le navigateur
    # -----------------------------------------------------
    return send_file(zip_path, as_attachment=True, download_name="resultats.zip")

# --------------------------------------------------------
# 4. Lancer l'application Flask en mode debug
# --------------------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)