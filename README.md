# Maltem Africa — CV Converter

Outil web + CLI qui transforme n'importe quel CV (PDF ou DOCX) en un CV reformaté au design officiel **Maltem Africa**, en utilisant l'IA **Kimi (NVIDIA)** pour extraire et structurer le contenu.

---

## 🚀 Démarrage rapide

### 1. Installer les dépendances

```bash
cd maltem-cv-converter
pip install -r requirements.txt
```

### 2. Configurer la clé API NVIDIA

```bash
cp .env.example .env
# Éditez .env et mettez votre clé NVIDIA_API_KEY
```

Ou directement dans le terminal :

```bash
# Linux / Mac
export NVIDIA_API_KEY="votre_clé_nvidia"

# Windows
set NVIDIA_API_KEY=votre_clé_nvidia
```

---

## 🌐 Interface Web

```bash
cd backend
python app.py
```

Ouvrez votre navigateur sur **http://localhost:5000**

1. Uploadez votre CV (PDF ou DOCX)
2. Cliquez sur **"Convertir au format Maltem"**
3. Le CV reformaté se télécharge automatiquement

---

## 💻 CLI (ligne de commande)

```bash
# Conversion simple
python cli/convert.py mon_cv.pdf

# Avec dossier de sortie personnalisé
python cli/convert.py mon_cv.docx --output ./resultats/

# Sauvegarder aussi les données JSON extraites
python cli/convert.py mon_cv.pdf --json donnees_extraites.json

# Mode verbose (affiche toutes les données extraites)
python cli/convert.py mon_cv.pdf --verbose
```

---

## 📁 Structure du projet

```
maltem-cv-converter/
├── backend/
│   ├── app.py              ← Serveur Flask (API + interface web)
│   ├── cv_parser.py        ← Extraction texte depuis PDF/DOCX
│   ├── kimi_extractor.py   ← Appel API Kimi NVIDIA
│   ├── cv_formatter.py     ← Génération DOCX style Maltem
│   ├── assets/
│   │   └── logo_maltem.png ← Logo officiel Maltem
│   ├── static/
│   │   └── index.html      ← Interface web
│   ├── uploads/            ← Fichiers uploadés (temporaires)
│   └── outputs/            ← CV générés
├── cli/
│   └── convert.py          ← CLI en ligne de commande
├── requirements.txt
├── .env.example
└── README.md
```

---

## 🔧 Flux de fonctionnement

```
CV utilisateur (PDF/DOCX)
        ↓
  [cv_parser.py]
  Extraction du texte brut
        ↓
  [kimi_extractor.py]
  API Kimi NVIDIA → JSON structuré
  (nom, poste, expériences, compétences...)
        ↓
  [cv_formatter.py]
  Génération DOCX — Design Maltem Africa
  (Century Gothic, rouge #E9272D, logo)
        ↓
  CV_Maltem_NomPrenom.docx ✓
```

---

## 📋 Dépendances

| Package | Rôle |
|---------|------|
| `flask` | Serveur web |
| `python-docx` | Lecture/écriture DOCX |
| `pdfplumber` | Extraction texte PDF |
| `requests` | Appels API NVIDIA |
| `werkzeug` | Gestion des uploads |

---

## ⚙️ Variables d'environnement

| Variable | Description | Défaut |
|----------|-------------|--------|
| `NVIDIA_API_KEY` | Clé API NVIDIA (obligatoire) | — |
| `PORT` | Port du serveur web | `5000` |

---

## 🎨 Design Maltem

Le CV généré respecte la charte graphique Maltem Africa :
- **Police** : Century Gothic
- **Couleur principale** : Rouge `#E9272D`
- **Logo** : Maltem Africa officiel
- **Sections** : À propos, Compétences, Certifications, Formation, Expériences, Projets marquants
