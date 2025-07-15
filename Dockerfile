# Utilise une image Python officielle légère
FROM python:3.11-slim

# Définit le répertoire de travail dans le conteneur
WORKDIR /app

# Copie les fichiers de ton projet
COPY . .

# Installe les dépendances
RUN pip install --no-cache-dir -r requirements.txt

# Expose le port 5000 pour Flask
EXPOSE 5000

# Commande pour lancer ton app
CMD ["python", "app.py"]
