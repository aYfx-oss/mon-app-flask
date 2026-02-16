"""
kimi_extractor.py — Extraction structurée du CV via Kimi K2 (NVIDIA NIM)
Approche multi-étapes optimisée pour gérer les CV longs
"""
import os, json, requests, re

NVIDIA_API_KEY = os.environ.get("NVIDIA_API_KEY", "")
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL = "moonshotai/kimi-k2-instruct"

def call_kimi(system_prompt: str, user_text: str, max_tokens: int = 8000) -> str:
    """Appel générique à Kimi avec timeout augmenté"""
    if not NVIDIA_API_KEY:
        raise ValueError("NVIDIA_API_KEY non défini")
    
    headers = {
        "Authorization": f"Bearer {NVIDIA_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_text}
        ],
        "temperature": 0.1,
        "max_tokens": max_tokens
    }
    
    # ✅ FIX 1: Augmenter le timeout à 240s (4 minutes)
    resp = requests.post(API_URL, headers=headers, json=payload, timeout=240)
    
    if resp.status_code != 200:
        raise RuntimeError(f"Erreur API Kimi: {resp.status_code} - {resp.text[:200]}")
    
    content = resp.json()["choices"][0]["message"]["content"].strip()
    
    # Nettoyer les balises markdown
    if content.startswith("```"):
        lines = content.split("\n")
        content = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
        content = content.replace("```json", "").replace("```", "").strip()
    
    return content


def extract_basic_info(text: str) -> dict:
    """Étape 1 : Extraire infos de base, compétences, langues, certifications"""
    # ✅ FIX 2: Limiter la taille du texte à 12000 caractères
    text_trimmed = text[:12000]
    
    prompt = """Extrais les informations de base du CV et retourne UNIQUEMENT un JSON valide :
{
  "nom_prenom": "Prénom NOM",
  "titre_poste": "Titre du poste actuel ou recherché",
  "annees_experience": "X ans d'expérience (calcule à partir des dates)",
  "a_propos": "Résumé/profil professionnel (ou chaîne vide)",
  "competences": [
    {"categorie": "Backend", "items": ["NodeJs", "Python"]},
    {"categorie": "Frontend", "items": ["React", "Vue.js"]}
  ],
  "certifications": ["Certification 1", "Certification 2"],
  "langues": ["Français", "Anglais"]
}
Retourne UNIQUEMENT le JSON, rien d'autre."""
    
    result = call_kimi(prompt, text_trimmed, max_tokens=3000)
    return json.loads(result)


def extract_experiences(text: str) -> list:
    """Étape 2 : Extraire les expériences professionnelles"""
    # ✅ FIX 2: Limiter la taille du texte à 15000 caractères
    text_trimmed = text[:15000]
    
    prompt = """Extrais TOUTES les expériences professionnelles du CV et retourne un JSON valide :
{
  "experiences": [
    {
      "periode": "2022 – 2024",
      "entreprise": "NOM ENTREPRISE",
      "poste": "Titre du poste",
      "direction": "Direction (si mentionné, sinon vide)",
      "contexte": "Contexte du projet/mission",
      "objectifs": ["objectif1", "objectif2"],
      "missions": ["mission1", "mission2"],
      "realisations": ["realisation1", "realisation2"],
      "resultats": ["résultat1", "résultat2"],
      "environnement": "Technologies, outils"
    }
  ]
}
Retourne UNIQUEMENT le JSON."""
    
    result = call_kimi(prompt, text_trimmed, max_tokens=10000)
    data = json.loads(result)
    return data.get("experiences", [])


def extract_formation_projets(text: str) -> dict:
    """Étape 3 : Extraire formation, projets, autres références"""
    # ✅ FIX 2: Limiter la taille du texte à 10000 caractères
    text_trimmed = text[:10000]
    
    prompt = """Extrais la formation, projets marquants et autres références du CV. Retourne un JSON valide :
{
  "formations": [
    {
      "annee": "2020",
      "diplome": "Nom du diplôme",
      "etablissement": "Établissement"
    }
  ],
  "projets_marquants": ["Projet 1", "Projet 2"],
  "autres_references": [
    {
      "entreprise": "Nom entreprise",
      "poste": "Poste"
    }
  ]
}
Retourne UNIQUEMENT le JSON."""
    
    result = call_kimi(prompt, text_trimmed, max_tokens=3000)
    return json.loads(result)


def extract_cv_data(text: str) -> dict:
    """
    Point d'entrée principal : traite le CV en 3 étapes
    pour gérer les CV longs sans dépasser la limite de contexte
    """
    print("🔄 Extraction CV en cours (multi-étapes optimisées)...")
    
    # Étape 1 : Infos de base
    print("  ➜ Étape 1/3 : Infos de base, compétences, langues...")
    cv_data = extract_basic_info(text)
    
    # Étape 2 : Expériences (peut être long)
    print("  ➜ Étape 2/3 : Expériences professionnelles...")
    cv_data["experiences"] = extract_experiences(text)
    
    # Étape 3 : Formation et projets
    print("  ➜ Étape 3/3 : Formation, projets, références...")
    extra = extract_formation_projets(text)
    cv_data.update(extra)
    
    print("✅ Extraction terminée !")
    return cv_data


# Alias pour compatibilité
structure_cv_with_kimi = extract_cv_data
