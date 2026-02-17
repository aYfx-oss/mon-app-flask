"""
kimi_extractor.py — Extraction structurée du CV via Kimi K2 (NVIDIA NIM)
Approche multi-étapes correcte : chaque appel reçoit peu de texte et génère peu de JSON
"""
import os, json, requests

NVIDIA_API_KEY = os.environ.get("NVIDIA_API_KEY", "")
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL   = "moonshotai/kimi-k2-instruct"


def call_kimi(prompt: str, text: str) -> str:
    """Appel générique à Kimi avec un texte limité"""
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

    # Nettoyer markdown
    if content.startswith("```"):
        lines = content.split("\n")
        content = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
        content = content.replace("```json", "").replace("```", "").strip()

    return content


def split_text(text: str, max_chars: int = 4000) -> list:
    """Découpe le texte en chunks"""
    chunks = []
    while len(text) > max_chars:
        split_at = text.rfind("\n", 0, max_chars)
        if split_at == -1:
            split_at = max_chars
        chunks.append(text[:split_at])
        text = text[split_at:]
    chunks.append(text)
    return chunks


def extract_cv_data(text: str) -> dict:
    """Extraction multi-étapes du CV"""

    chunks = split_text(text, max_chars=4000)
    first_chunk = chunks[0]
    all_text = text[:8000]  # Pour les expériences on prend plus

    # ── Étape 1 : Infos de base ──────────────────────────────────────────────
    prompt1 = """Extrais uniquement ces informations du CV et retourne ce JSON :
{
  "nom_prenom": "Prénom NOM",
  "titre_poste": "Titre du poste",
  "annees_experience": "X ans d'expérience",
  "a_propos": "Si résumé présent → copie-le. Sinon → génère 2-3 phrases pro basées sur le profil",
  "competences": [{"categorie": "Catégorie", "items": ["item1", "item2"]}],
  "certifications": ["cert1", "cert2"],
  "langues": ["Français", "Anglais"]
}"""

    result1 = json.loads(call_kimi(prompt1, first_chunk))

    # ── Étape 2 : Expériences ────────────────────────────────────────────────
    prompt2 = """Extrais TOUTES les expériences professionnelles et retourne ce JSON :
{
  "experiences": [
    {
      "periode": "2022 – 2024",
      "entreprise": "ENTREPRISE",
      "poste": "Poste",
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
    prompt3 = """Extrais la formation, projets marquants et autres références et retourne ce JSON :
{
  "formations": [{"annee": "2020", "diplome": "Diplôme", "etablissement": "École"}],
  "projets_marquants": ["projet1", "projet2"],
  "autres_references": [{"entreprise": "Entreprise", "poste": "Poste"}]
}"""

    result3 = json.loads(call_kimi(prompt3, first_chunk))

    # ── Fusion des résultats ─────────────────────────────────────────────────
    cv_data = {}
    cv_data.update(result1)
    cv_data.update(result2)
    cv_data.update(result3)

    return cv_data


structure_cv_with_kimi = extract_cv_data
