"""
kimi_extractor.py — Extraction structurée du CV via Kimi K2 (NVIDIA NIM)
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
  "competences": [{"categorie": "Catégorie", "items": ["item1"]}],
  "certifications": ["cert1"],
  "formations": [{"annee": "2020", "diplome": "Diplôme", "etablissement": "École"}],
  "experiences": [
    {
      "periode": "2022 – 2024",
      "entreprise": "ENTREPRISE",
      "poste": "Poste",
      "direction": "",
      "contexte": "Contexte",
      "objectifs": [],
      "missions": ["mission1"],
      "realisations": ["realisation1"],
      "resultats": [],
      "environnement": "Technologies"
    }
  ],
  "autres_references": [],
  "projets_marquants": [],
  "langues": ["Français"]
}

Règles :
- annees_experience : calcule depuis les dates, écris "X ans d'expérience"
- a_propos : SI le CV a un résumé → copie-le. SINON → génère 2-3 phrases professionnelles basées sur ses expériences et compétences
- Champs absents : liste vide [] ou chaîne vide ""
- UNIQUEMENT le JSON, rien d'autre"""


def extract_cv_data(text: str) -> dict:
    if not NVIDIA_API_KEY:
        raise ValueError("NVIDIA_API_KEY non défini")

    # Tronquer le texte si trop long (max ~6000 caractères)
    if len(text) > 6000:
        text = text[:6000] + "\n...[texte tronqué]"

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

    return json.loads(content)

structure_cv_with_kimi = extract_cv_data
