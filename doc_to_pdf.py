import os
import subprocess
import logging

# Configuration de logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Chemin vers le dossier 'documents'
output_dir = 'documents'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Liste pour stocker les chemins des fichiers DOCX générés
docx_files = [os.path.join(output_dir, f) for f in os.listdir(output_dir) if f.endswith('.docx')]
pdf_files = []  # Liste pour stocker les chemins des fichiers PDF générés

# Fonction pour convertir DOCX en PDF
def convert_docx_to_pdf(docx_path):
    # Ajuste bien le chemin d'installation de LibreOffice si nécessaire
    libreoffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    cmd = [libreoffice_path, '--convert-to', 'pdf', '--outdir', output_dir, docx_path, '--headless']
    subprocess.run(cmd, check=True)
    pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
    if os.path.exists(pdf_path):
        logging.debug(f"Fichier PDF créé : {pdf_path}")
        pdf_files.append(pdf_path)
    else:
        logging.error(f"Le fichier PDF attendu n'a pas été créé : {pdf_path}")

# Convertir chaque DOCX en PDF
for docx_file in docx_files:
    convert_docx_to_pdf(docx_file)

# Afficher l'ordre des fichiers PDF avant fusion
logging.debug("Ordre des fichiers PDF avant fusion : " + str(pdf_files))

# Fonction de fusion des PDF utilisant pdftk
def merge_pdfs(pdf_files, output_pdf):
    # Trier les fichiers par ordre alphabétique
    pdf_files.sort()
    logging.debug("Ordre des fichiers PDF après tri : " + str(pdf_files))
    cmd = ["pdftk"] + pdf_files + ["cat", "output", output_pdf]
    subprocess.run(cmd, check=True)

# ==============================
#  FUSION DES PDF SI PLUSIEURS
# ==============================
if len(pdf_files) > 1:
    output_pdf_path = os.path.join(output_dir, 'Final_Merged_Document.pdf')
    try:
        merge_pdfs(pdf_files, output_pdf_path)
        logging.info(f"Final merged PDF created: {output_pdf_path}")
    except Exception as e:
        logging.error(f"Failed to merge PDFs. Error: {e}")
else:
    logging.info("No multiple PDFs to merge or only one PDF file exists.")