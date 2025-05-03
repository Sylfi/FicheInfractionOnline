# utils/commune.py
import requests
import logging

def fetch_commune_code(commune_name, postal_code):
    """
    Trouve le code INSEE d'une commune en utilisant son nom et son code postal.
    :param commune_name: Nom de la commune
    :param postal_code: Code postal de la commune
    :return: Code INSEE de la commune ou None si non trouvé
    """
    url = f"https://geo.api.gouv.fr/communes?nom={commune_name}&limit=100"
    try:
        response = requests.get(url)
        response.raise_for_status()  # Vérifie si la requête a réussi
        data = response.json()  # Conversion du résultat en JSON

        # Filtrer les communes par code postal pour s'assurer de la bonne correspondance
        filtered_communes = [commune for commune in data if postal_code in commune['codesPostaux']]

        if len(filtered_communes) == 1:
            return filtered_communes[0]['code']
        elif len(filtered_communes) > 1:
            logging.info(f"Plusieurs communes trouvées pour '{commune_name}' avec le code postal {postal_code}, veuillez préciser.")
            return None
        else:
            logging.info(f"Aucune commune trouvée pour '{commune_name}' avec le code postal {postal_code}.")
            return None
    except requests.exceptions.RequestException as e:
        logging.error(f"Erreur lors de la recherche des communes : {e}")
        return None