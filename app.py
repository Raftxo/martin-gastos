"""
app.py — Servidor Flask para la web app de gastos.
Ejecutar con: python app.py
"""

import os
import json
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from fill_excel_v9 import fill_excel, parse_csv_for_unknowns, convert_to_pdf

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB máximo

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "excel_templates")
OUTPUTS_DIR   = os.path.join(BASE_DIR, "outputs")
UPLOADS_DIR   = os.path.join(BASE_DIR, "uploads")

for d in [TEMPLATES_DIR, OUTPUTS_DIR, UPLOADS_DIR]:
    os.makedirs(d, exist_ok=True)


# ── Rutas ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/templates")
def list_templates():
    files = [f for f in os.listdir(TEMPLATES_DIR) if f.endswith(".xlsx")]
    # Poner la plantilla Raftxo primera si existe
    files.sort(key=lambda f: (0 if "Raftxo" in f else 1, f))
    return jsonify(files)


@app.route("/api/analyze", methods=["POST"])
def analyze():
    """
    Recibe el CSV, lo analiza y devuelve:
      - regiones desconocidas que necesitan ciudad
    """
    if "csv" not in request.files:
        return jsonify({"error": "No se recibió fichero CSV"}), 400

    csv_file = request.files["csv"]
    filename = secure_filename(csv_file.filename or "upload")
    csv_path = os.path.join(UPLOADS_DIR, filename)
    csv_file.save(csv_path)

    try:
        unknown_regions = parse_csv_for_unknowns(csv_path)
        return jsonify({
            "csv_path": csv_path,
            "unknown_regions": list(unknown_regions)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def generate():
    """
    Recibe el JSON con toda la información del formulario y genera el Excel.
    Body esperado:
    {
        "csv_path":       "ruta guardada en /analyze",
        "template":       "nombre_plantilla.xlsx",
        "location_map":   {"Andalucía": "Sevilla (AN)", ...},
        "manual_shifts":  [ { fecha, h_ini, h_fin, concepto, km } , ... ]
    }
    """
    data = request.get_json()

    csv_path      = data.get("csv_path")
    template_name = data.get("template")
    location_map  = data.get("location_map", {})
    manual_shifts = data.get("manual_shifts", [])

    if not csv_path or not template_name:
        return jsonify({"error": "Faltan csv_path o template"}), 400

    template_path = os.path.join(TEMPLATES_DIR, template_name)
    if not os.path.exists(template_path):
        return jsonify({"error": f"Plantilla no encontrada: {template_name}"}), 404

    try:
        output_filename = fill_excel(
            excel_path    = template_path,
            csv_path      = csv_path,
            location_map  = location_map,
            manual_shifts = manual_shifts,
            output_dir    = OUTPUTS_DIR,
        )
        # Generar PDF
        xlsx_abs = os.path.join(OUTPUTS_DIR, output_filename)
        convert_to_pdf(xlsx_abs)

        return jsonify({"output": output_filename})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/outputs/<filename>")
def serve_output(filename):
    mimetype = 'application/pdf' if filename.endswith('.pdf') else None
    return send_from_directory(OUTPUTS_DIR, filename, mimetype=mimetype)


# ── Arranque ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(debug=True, port=5000)
