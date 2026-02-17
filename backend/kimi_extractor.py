"""
kimi_extractor.py — Extraction structurée du CV via Kimi K2 (NVIDIA NIM)
Structure mise à jour pour le design ZAID/Maltem Africa.
"""
import os, json, requests

NVIDIA_API_KEY = os.environ.get("NVIDIA_API_KEY", "")
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL   = "moonshotai/kimi-k2-instruct"

SYSTEM_PROMPT = """Tu es un expert en extraction de données de CV professionnels.
Extrais les informations du CV fourni et retourne UNIQUEMENT un JSON valide, sans texte avant ou après, sans balises markdown.

Structure JSON exacte à retourner :
{
  "nom_prenom": "Prénom NOM",
  "titre_poste": "Titre du poste",
  "annees_experience": "X ans d'expérience",
  "a_propos": "Texte de présentation",
  "competences": [
    {
      "categorie": "Nom de la catégorie",
      "items": ["item1", "item2"]
    }
  ],
  "certifications": ["cert1", "cert2"],
  "formations": [
    {
      "annee": "2020",
      "diplome": "Nom du diplôme",
      "etablissement": "Nom de l'établissement"
    }
  ],
  "experiences": [
    {
      "periode": "2022 – 2024",
      "entreprise": "NOM ENTREPRISE",
      "poste": "Titre du poste",
      "direction": "Direction ou département (si mentionné, sinon chaîne vide)",
      "contexte": "Description du contexte du projet/mission",
      "objectifs": ["objectif1", "objectif2"],
      "missions": ["mission1", "mission2"],
      "realisations": ["realisation1", "realisation2"],
      "resultats": ["résultat1", "résultat2"],
      "environnement": "Technologies, outils, méthodologies utilisés"
    }
  ],
  "autres_references": [
    {
      "entreprise": "Nom entreprise",
      "poste": "Poste occupé"
    }
  ],
  "projets_marquants": ["projet1", "projet2"],
  "langues": ["Français", "Anglais"]
}

Règles importantes :
- Pour annees_experience : calcule à partir des dates et écris "X ans d'expérience"
- Pour a_propos : 
  * SI le CV contient un résumé/profil/à propos écrit par le candidat → copie-le tel quel
  * SI le CV ne contient PAS de résumé/profil → génère un paragraphe professionnel de 2-3 phrases basé sur l'analyse de ses expériences, compétences et formations. Ce texte doit refléter fidèlement son profil réel.
- Pour les expériences : remplis objectifs, missions, realisations selon ce qui est dans le CV
- Si un champ n'existe pas dans le CV, utilise une liste vide [] ou une chaîne vide ""
- Retourne UNIQUEMENT le JSON, pas d'explication"""


def extract_cv_data(text: str) -> dict:
    if not NVIDIA_API_KEY:
        raise ValueError("NVIDIA_API_KEY non défini")

    headers = {
        "Authorization": f"Bearer {NVIDIA_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": f"Extrais les données de ce CV :\n\n{text}"}
        ],
        "temperature": 0.1,
        "max_tokens": 16000
    }

    resp = requests.post(API_URL, headers=headers, json=payload, timeout=180)

    if resp.status_code != 200:
        raise ValueError(f"Erreur API Kimi: {resp.status_code} - {resp.text[:200]}")

    content = resp.json()["choices"][0]["message"]["content"].strip()

    # Nettoyer les balises markdown si présentes
    if content.startswith("```"):
        lines = content.split("\n")
        content = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])

    return json.loads(content)

structure_cv_with_kimi = extract_cv_data
