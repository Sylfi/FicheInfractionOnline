#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Module de génération des lettres recommandées et du fichier destinataires
pour les infractions d'une commune.
"""

import os
import sys
import csv
import unicodedata
import requests
import logging
from pathlib import Path
from docxtpl import DocxTemplate
import locale
from datetime import datetime
from typing import Optional


# Réglage de la locale pour les dates
locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

# URLs des API publiques
GEO_URL = "https://geo.api.gouv.fr/communes"
MAIRIE_URL = "https://etablissements-publics.api.gouv.fr/v3/communes/{insee}/mairie"


def strip_accents(s: str) -> str:
    """Supprime les accents et passe en minuscules pour comparer."""
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()


def find_commune(ville: str, cp: str) -> dict:
    """Retourne le dictionnaire de la commune via l’API geo.api.gouv.fr."""
    params = {
        "nom": ville,
        "codePostal": cp,
        "fields": "nom,code,codesPostaux,population",
        "format": "json",
    }
    resp = requests.get(GEO_URL, params=params, timeout=4)
    resp.raise_for_status()
    communes = resp.json()
    if not communes:
        sys.exit(f"Aucune commune '{ville}' (CP {cp}) trouvée.")
    # Filtrer sur le nom exact (accents ignorés)
    exact = [c for c in communes if strip_accents(c["nom"]) == strip_accents(ville)]
    choix = exact or communes
    return max(choix, key=lambda c: c.get("population", 0))


def get_mairie_address(insee: str) -> str:
    """Récupère l’adresse de la mairie via l’API établissements-publics."""
    resp = requests.get(MAIRIE_URL.format(insee=insee), timeout=4)
    resp.raise_for_status()
    features = resp.json().get("features", [])
    if not features:
        return "Adresse non disponible"
    prop = features[0]["properties"]
    adresses = prop.get("adresses", [])
    if adresses:
        principale = next((a for a in adresses if a.get("type") == "Adresse"), adresses[0])
        lignes = principale.get("lignes", [])
        return ", ".join(lignes) if lignes else "Adresse non disponible"
    return "Adresse non disponible"


def get_mayor_name_from_csv(insee_code: str, utils_dir: str) -> Optional[str]:
    """
    Cherche le maire dans utils/RNE.csv à partir du code INSEE.
    """
    csv_file = Path(utils_dir) / "RNE.csv"
    if not csv_file.exists():
        logging.error(f"RNE.csv introuvable : {csv_file}")
        return None
    try:
        with csv_file.open(newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                if row.get("Code de la commune") == insee_code:
                    prenom = row.get("Prénom de l'élu", "").title()
                    nom = row.get("Nom de l'élu", "").upper()
                    sexe = row.get("Code sexe", "").upper()
                    if sexe not in ("M", "F"):
                        sexe = "?"
                    return f"{prenom} {nom} ({sexe})".strip()
    except Exception as e:
        logging.error(f"Erreur lecture RNE.csv : {e}")
    return None


def generate_courrier(fiches: list[dict], utils_dir: str, courriers_dir: str):
    """
    Génère dans `courriers_dir` pour la commune représentée par `fiches` :
      - un fichier Lettre_Infractions_<Commune>.docx
      - un fichier destinataires.txt
    """
    if not fiches:
        logging.warning("Aucune fiche transmise au module de courrier.")
        return

    # Extraire la commune et CP depuis la première fiche
    ville = fiches[0]["Ville"]
    cp = fiches[0]["Code postal"]

    # Recherche INSEE et mairie
    commune = find_commune(ville, cp)
    insee = commune["code"]
    mairie_address = get_mairie_address(insee)

    # Recherche du maire
    mayor = get_mayor_name_from_csv(insee, utils_dir) or ""
    prenom_maire, nom_maire, sexe_code = "", "", "M"
    parts = mayor.split()
    if len(parts) >= 2:
        prenom_maire = parts[0]
        if "(" in parts[-1]:
            # dernier élément "(M)" ou "(F)"
            sexe = parts[-1].strip("()")
            sexe_code = sexe if sexe in ("M", "F") else "M"
            nom_maire = " ".join(parts[1:-1]) or parts[1]
        else:
            nom_maire = " ".join(parts[1:])

    # Comptage des fiches
    nombre_de_fiches = len(fiches)
    code_fiche_01 = fiches[0]["Nom"]
    code_fiche_derniere = fiches[-1]["Nom"]

    # Préparer dossier de sortie
    os.makedirs(courriers_dir, exist_ok=True)

    # Chargement du modèle de lettre
    letter_template = os.path.join(utils_dir, "modele_lettre_infraction.docx")
    if not os.path.isfile(letter_template):
        logging.error(f"Template lettre introuvable : {letter_template}")
        return
    lettre_doc = DocxTemplate(letter_template)

    # Contexte pour la lettre
    date_str = datetime.now().strftime("%d %B %Y")
    context = {
        "date_today": date_str,
        "pronom_maire": "Madame" if sexe_code == "F" else "Monsieur",
        "le_la": "la" if sexe_code == "F" else "le",
        "prenom_maire": prenom_maire,
        "nom_maire": nom_maire,
        "nom_commune": ville,
        "adresse_mairie": mairie_address,
        "code_postal": cp,
        "seul_e_competent_e": "seule compétente" if sexe_code == "F" else "seul compétent",
        "nombre_de_fiches": nombre_de_fiches,
        "code_fiche_01": code_fiche_01,
        "code_fiche_dernier_numero": code_fiche_derniere,
    }
    lettre_doc.render(context)

    # Sauvegarde de la lettre
    date_iso = datetime.now().strftime("%Y-%m-%d")
    lettre_path = os.path.join(courriers_dir, f"{date_iso}-demande initiale.docx")
    lettre_doc.save(lettre_path)
    logging.info(f"Courrier généré : {lettre_path}")

    # Génération du fichier destinataires.txt
    dest_cols = [
        "Société", "Civilité", "Prénom", "Nom",
        "Batiment", "Libellé voie", "Code postal", "Ville", "Pays"
    ]
    civ = "Madame" if sexe_code == "F" else "Monsieur"
    soc = f"{civ} {prenom_maire} {nom_maire}".strip()
    civ_nom = f"Maire de {ville}"
    dest_vals = {
        "Société": soc,
        "Civilité": civ_nom,
        "Prénom": civ_nom,
        "Nom": civ_nom,
        "Batiment": "Hôtel de Ville",
        "Libellé voie": mairie_address,
        "Code postal": f"{cp} {ville}",
        "Ville": "",
        "Pays": "France"
    }
    destinataires_path = os.path.join(courriers_dir, "destinataires.txt")
    with open(destinataires_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=dest_cols, delimiter="\t")
        writer.writeheader()
        writer.writerow(dest_vals)
    logging.info(f"Fichier destinataires généré : {destinataires_path}")