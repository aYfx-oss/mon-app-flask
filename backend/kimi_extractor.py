"""
kimi_extractor.py — Extraction structurée du CV via Kimi K2 (NVIDIA NIM)
"""
import os, json, requests

NVIDIA_API_KEY = os.environ.get("NVIDIA_API_KEY", "")
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL   = "moonshotai/kimi-k2-instruct"


def call_kimi(prompt: str, text: str) -> str:
    if not NVIDIA_API_KEY:
        raise ValueError("NVIDIA_API_KEY non défini")

    headers = {
        "Authorization": f"Bearer {NVIDIA_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": "Tu es un expert en extraction de données de CV. Retourne UNIQUEMENT du JSON valide, sans markdown, sans explication."},
            {"role": "user", "content": f"{prompt}\n\nCV:\n{text}"}
        ],
        "temperature": 0.1,
        "max_tokens": 4096
    }

    resp = requests.post(API_URL, headers=headers, json=payload, timeout=180)

    if resp.status_code != 200:
        raise RuntimeError(f"Erreur API Kimi: {resp.status_code} - {resp.text[:200]}")

    content = resp.json()["choices"][0]["message"]["content"].strip()

    if content.startswith("```"):
        lines = content.split("\n")
        content = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
        content = content.replace("```json", "").replace("```", "").strip()

    return content


def extract_cv_data(text: str) -> dict:

    first_chunk = text[:4000]
    all_text = text[:8000]

    # ── Étape 1 : Infos de base ──────────────────────────────────────────────
    prompt1 = """Analyse attentivement ce CV et extrais ces informations.
IMPORTANT : Le nom et prénom sont généralement au tout début du CV, souvent en titre ou en gros.
Cherche bien le vrai nom de la personne, ne mets JAMAIS "Prénom NOM" comme valeur.

Retourne ce JSON exact :
{
  "nom_prenom": "NOM ET PRÉNOM",
  "titre_poste": "Titre du poste actuel ou recherché",
  "annees_experience": "X ans d'expérience (calcule depuis les dates)",
  "a_propos": "Si résumé présent dans le CV → copie-le exactement. Sinon → génère 2-3 phrases professionnelles basées sur le profil réel de la personne",
  "competences": [{"categorie": "Catégorie", "items": ["item1", "item2"]}],
  "certifications": ["cert1", "cert2"],
  "langues": ["Français", "Anglais"]
}"""

    result1 = json.loads(call_kimi(prompt1, first_chunk))

    # ── Étape 2 : Expériences ────────────────────────────────────────────────
    prompt2 = """Extrais TOUTES les expériences professionnelles du CV et retourne ce JSON :
{
  "experiences": [
    {
      "periode": "2022 – 2024",
      "entreprise": "NOM ENTREPRISE",
      "poste": "Titre du poste",
      "direction": "",
      "contexte": "Contexte de la mission",
      "objectifs": [],
      "missions": ["mission1", "mission2"],
      "realisations": ["realisation1"],
      "resultats": [],
      "environnement": "Technologies utilisées"
    }
  ]
}"""

    result2 = json.loads(call_kimi(prompt2, all_text))

    # ── Étape 3 : Formation et projets ───────────────────────────────────────
    prompt3 = """Extrais la formation, projets marquants et autres références du CV et retourne ce JSON :
{
  "formations": [{"annee": "2020", "diplome": "Diplôme", "etablissement": "École"}],
  "projets_marquants": ["projet1", "projet2"],
  "autres_references": [{"entreprise": "Entreprise", "poste": "Poste"}]
}"""

    result3 = json.loads(call_kimi(prompt3, first_chunk))

    # ── Fusion ───────────────────────────────────────────────────────────────
    cv_data = {}
    cv_data.update(result1)
    cv_data.update(result2)
    cv_data.update(result3)

    return cv_data


structure_cv_with_kimi = extract_cv_data
