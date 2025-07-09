# Utilise une image de base avec Python et Node.js
FROM python:3.11-slim

# Installe Node.js pour Mermaid CLI
RUN apt-get update && apt-get install -y \
    curl \
    gnupg \
    && curl -fsSL https://deb.nodesource.com/setup_18.x | bash - \
    && apt-get install -y nodejs \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Installe Mermaid CLI globalement
RUN npm install -g @mermaid-js/mermaid-cli@latest

# Définit le répertoire de travail
WORKDIR /app

# Copie les fichiers de requirements en premier (pour optimiser le cache Docker)
COPY requirements.txt .

# Installe les dépendances Python
RUN pip install --no-cache-dir -r requirements.txt

# Copie tout le code de l'application
COPY . .

# Crée les dossiers nécessaires pour les uploads et outputs
RUN mkdir -p uploads outputs

# Expose le port 5000
EXPOSE 5000

# Variables d'environnement
ENV FLASK_APP=app.py
ENV FLASK_ENV=production
ENV PYTHONPATH=/app

# Vérifie que mmdc est disponible
RUN mmdc --version

# Commande par défaut pour lancer l'application
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "app:app"]
