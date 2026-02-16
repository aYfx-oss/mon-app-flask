"""
app.py — Serveur web Flask pour le convertisseur de CV Maltem Africa
"""
import os
import uuid
from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename

from cv_parser import extract_cv_text
from kimi_extractor import structure_cv_with_kimi
from cv_formatter import generate_maltem_cv

# ── Configuration ────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="static")

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), "outputs")
ALLOWED_EXTENSIONS = {"pdf", "docx"}
MAX_CONTENT_LENGTH = 10 * 1024 * 1024  # 10 MB max

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/convert", methods=["POST"])
def convert_cv():
    """
    Endpoint principal :
    1. Reçoit le fichier CV (PDF ou DOCX)
    2. Extrait le texte
    3. Structurise avec Kimi NVIDIA
    4. Génère le CV Maltem DOCX
    5. Retourne le fichier
    """
    # Vérification du fichier
    if "cv_file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu. Utilisez le champ 'cv_file'."}), 400

    file = request.files["cv_file"]

    if file.filename == "":
        return jsonify({"error": "Nom de fichier vide."}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Format non supporté. Envoyez un fichier PDF ou DOCX."}), 400

    # Sauvegarder le fichier uploadé
    unique_id = str(uuid.uuid4())[:8]
    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, f"{unique_id}_{filename}")
    file.save(input_path)

    try:
        # Étape 1 : Extraction du texte
        raw_text = extract_cv_text(input_path)
        if not raw_text.strip():
            return jsonify({"error": "Impossible d'extraire le texte du CV. Le fichier semble vide ou protégé."}), 400

        # Étape 2 : Structuration avec Kimi NVIDIA
        cv_data = structure_cv_with_kimi(raw_text)

        # Étape 3 : Génération du CV Maltem
        nom = cv_data.get("nom_prenom", "CV").replace(" ", "_")
        output_filename = f"CV_Maltem_{nom}_{unique_id}.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        generate_maltem_cv(cv_data, output_path)

        # Étape 4 : Envoi du fichier
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 502
    except Exception as e:
        return jsonify({"error": f"Erreur interne : {str(e)}"}), 500
    finally:
        # Nettoyer le fichier uploadé
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "Maltem CV Converter"})


# ── Lancement ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"🚀 Maltem CV Converter démarré sur http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
