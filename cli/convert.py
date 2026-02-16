#!/usr/bin/env python3
"""
cli/convert.py — CLI Maltem CV Converter
Utilisation : python convert.py <chemin_cv> [--output <dossier_sortie>]

Exemple :
  python convert.py mon_cv.pdf
  python convert.py mon_cv.docx --output ./output/
"""

import os
import sys
import argparse
import json

# Ajouter le backend au path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from cv_parser import extract_cv_text
from kimi_extractor import structure_cv_with_kimi
from cv_formatter import generate_maltem_cv


# ── Couleurs ANSI pour le terminal ────────────────────────────────────────────
RED    = "\033[91m"
GREEN  = "\033[92m"
YELLOW = "\033[93m"
BLUE   = "\033[94m"
BOLD   = "\033[1m"
RESET  = "\033[0m"


def banner():
    print(f"""
{RED}{BOLD}╔══════════════════════════════════════════╗
║     MALTEM AFRICA — CV CONVERTER CLI     ║
║          Powered by Kimi AI (NVIDIA)     ║
╚══════════════════════════════════════════╝{RESET}
""")


def step(n, total, msg):
    print(f"{BLUE}[{n}/{total}]{RESET} {msg}...")


def success(msg):
    print(f"{GREEN}✓{RESET} {msg}")


def error(msg):
    print(f"{RED}✗ ERREUR:{RESET} {msg}", file=sys.stderr)


def check_env():
    """Vérifie que la clé API NVIDIA est configurée."""
    key = os.environ.get("NVIDIA_API_KEY", "")
    if not key:
        error("La variable d'environnement NVIDIA_API_KEY n'est pas définie.")
        print(f"\n{YELLOW}Configurez-la avec :{RESET}")
        print("  export NVIDIA_API_KEY='votre_clé_nvidia'   (Linux/Mac)")
        print("  set NVIDIA_API_KEY=votre_clé_nvidia         (Windows)")
        sys.exit(1)


def main():
    banner()

    parser = argparse.ArgumentParser(
        description="Convertit un CV (PDF/DOCX) au format Maltem Africa",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemples :
  python convert.py mon_cv.pdf
  python convert.py mon_cv.docx --output ./resultats/
  python convert.py mon_cv.pdf --json cv_extrait.json
        """
    )
    parser.add_argument("input", help="Chemin vers le CV source (PDF ou DOCX)")
    parser.add_argument("--output", "-o", default=".", help="Dossier de sortie (défaut : répertoire courant)")
    parser.add_argument("--json", "-j", help="Sauvegarder les données extraites en JSON (optionnel)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Afficher les données extraites")

    args = parser.parse_args()

    # ── Vérifications ────────────────────────────────────────────────────────
    check_env()

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        error(f"Fichier introuvable : {input_path}")
        sys.exit(1)

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in (".pdf", ".docx"):
        error(f"Format non supporté : {ext}. Utilisez PDF ou DOCX.")
        sys.exit(1)

    output_dir = os.path.abspath(args.output)
    os.makedirs(output_dir, exist_ok=True)

    print(f"{BOLD}Fichier source :{RESET} {input_path}")
    print(f"{BOLD}Dossier sortie :{RESET} {output_dir}\n")

    TOTAL_STEPS = 3

    # ── Étape 1 : Extraction ─────────────────────────────────────────────────
    step(1, TOTAL_STEPS, "Extraction du texte du CV")
    try:
        raw_text = extract_cv_text(input_path)
        if not raw_text.strip():
            error("Aucun texte extrait. Le fichier semble vide ou protégé.")
            sys.exit(1)
        success(f"Texte extrait ({len(raw_text)} caractères)")
    except Exception as e:
        error(f"Échec de l'extraction : {e}")
        sys.exit(1)

    # ── Étape 2 : Analyse IA ─────────────────────────────────────────────────
    step(2, TOTAL_STEPS, "Analyse et structuration avec Kimi AI (NVIDIA)")
    try:
        cv_data = structure_cv_with_kimi(raw_text)
        nom = cv_data.get("nom_prenom", "—")
        poste = cv_data.get("titre_poste", "—")
        success(f"CV analysé : {nom} | {poste}")
    except Exception as e:
        error(f"Échec de l'analyse Kimi : {e}")
        sys.exit(1)

    # Affichage verbose
    if args.verbose:
        print(f"\n{YELLOW}── Données extraites ──{RESET}")
        print(json.dumps(cv_data, ensure_ascii=False, indent=2))
        print()

    # Sauvegarder en JSON si demandé
    if args.json:
        json_path = os.path.abspath(args.json)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(cv_data, f, ensure_ascii=False, indent=2)
        print(f"  {YELLOW}JSON sauvegardé :{RESET} {json_path}")

    # ── Étape 3 : Génération DOCX ────────────────────────────────────────────
    step(3, TOTAL_STEPS, "Génération du CV au format Maltem Africa")
    try:
        nom_clean = cv_data.get("nom_prenom", "CV").replace(" ", "_").replace("/", "_")
        output_filename = f"CV_Maltem_{nom_clean}.docx"
        output_path = os.path.join(output_dir, output_filename)
        generate_maltem_cv(cv_data, output_path)
        success(f"CV généré avec succès !")
    except Exception as e:
        error(f"Échec de la génération DOCX : {e}")
        sys.exit(1)

    # ── Résultat ─────────────────────────────────────────────────────────────
    print(f"""
{GREEN}{BOLD}╔══════════════════════════════════════════╗
║            CONVERSION RÉUSSIE !          ║
╚══════════════════════════════════════════╝{RESET}

{BOLD}Fichier généré :{RESET}
  {GREEN}{output_path}{RESET}
""")


if __name__ == "__main__":
    main()
