"""
app.py — Servidor Flask para la web app de gastos.
Ejecutar con: python app.py
"""

import os
import json
import logging
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from fill_excel_v9 import fill_excel, parse_csv_for_unknowns, convert_to_pdf, export_preview_png

# ── Logging ────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

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
    logger.info("Received /api/analyze request")
    
    if "csv" not in request.files:
        logger.warning("No CSV file received in request")
        return jsonify({"error": "No se recibió fichero CSV"}), 400

    csv_file = request.files["csv"]
    filename = secure_filename(csv_file.filename or "upload")
    csv_path = os.path.join(UPLOADS_DIR, filename)
    
    logger.info(f"Saving uploaded CSV: {filename} to {csv_path}")
    csv_file.save(csv_path)

    try:
        logger.info(f"Analyzing CSV for unknown regions: {csv_path}")
        unknown_regions = parse_csv_for_unknowns(csv_path)
        logger.info(f"Found {len(unknown_regions)} unknown regions: {unknown_regions}")
        
        return jsonify({
            "csv_path": csv_path,
            "unknown_regions": list(unknown_regions)
        })
    except ValueError as e:
        logger.error(f"CSV validation error: {e}")
        return jsonify({"error": f"CSV validation failed: {str(e)}"}), 400
    except Exception as e:
        logger.exception(f"Unexpected error in /api/analyze: {e}")
        return jsonify({"error": f"Error procesando CSV: {str(e)}"}), 500


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
    logger.info("Received /api/generate request")
    
    data = request.get_json()
    csv_path      = data.get("csv_path")
    template_name = data.get("template")
    location_map  = data.get("location_map", {})
    manual_shifts = data.get("manual_shifts", [])

    # Validate inputs
    if not csv_path or not template_name:
        logger.warning(f"Missing required fields: csv_path={csv_path}, template={template_name}")
        return jsonify({"error": "Faltan csv_path o template"}), 400

    # Validate manual_shifts limit
    if len(manual_shifts) > 50:
        logger.warning(f"Too many manual shifts: {len(manual_shifts)} > 50")
        return jsonify({"error": f"Máximo 50 actividades manuales permitidas ({len(manual_shifts)} recibidas)"}), 400

    # Validate CSV file exists
    if not os.path.exists(csv_path):
        logger.error(f"CSV file not found: {csv_path}")
        return jsonify({"error": f"Fichero CSV no encontrado: {csv_path}"}), 404

    template_path = os.path.join(TEMPLATES_DIR, template_name)
    if not os.path.exists(template_path):
        logger.error(f"Template not found: {template_path}")
        return jsonify({"error": f"Plantilla no encontrada: {template_name}"}), 404

    logger.info(f"Generating Excel with {len(manual_shifts)} manual shifts")

    try:
        output_filename = fill_excel(
            excel_path    = template_path,
            csv_path      = csv_path,
            location_map  = location_map,
            manual_shifts = manual_shifts,
            output_dir    = OUTPUTS_DIR,
        )
        logger.info(f"Excel generated successfully: {output_filename}")
        
        # Generar PDF
        xlsx_abs = os.path.join(OUTPUTS_DIR, output_filename)
        if not os.path.exists(xlsx_abs):
            logger.error(f"Generated XLSX file not found: {xlsx_abs}")
            return jsonify({"error": f"Error: Fichero XLSX no generado correctamente"}), 500
        
        logger.info(f"Converting to PDF: {xlsx_abs}")
        try:
            convert_to_pdf(xlsx_abs)
            logger.info("PDF conversion successful")
        except Exception as pdf_err:
            logger.warning(f"PDF conversion failed (non-fatal): {pdf_err}")
            # Don't fail the request if PDF fails — XLSX is still valid

        preview_filename = None
        logger.info(f"Generating PNG preview: {xlsx_abs}")
        try:
            preview_filename = export_preview_png(xlsx_abs)
            logger.info(f"PNG preview successful: {preview_filename}")
        except Exception as preview_err:
            logger.warning(f"PNG preview failed (non-fatal): {preview_err}")
        
        # Clean up the uploaded CSV file
        try:
            if os.path.exists(csv_path):
                os.remove(csv_path)
                logger.info(f"Cleaned up uploaded CSV: {csv_path}")
        except Exception as cleanup_err:
            logger.warning(f"Failed to cleanup CSV file: {cleanup_err}")
        
        return jsonify({"output": output_filename, "preview": preview_filename})
    except ValueError as e:
        logger.error(f"Validation error in Excel generation: {e}")
        return jsonify({"error": f"Error de validación: {str(e)}"}), 400
    except Exception as e:
        logger.exception(f"Unexpected error in /api/generate: {e}")
        return jsonify({"error": f"Error generando Excel: {str(e)}"}), 500


@app.route("/outputs/<filename>")
def serve_output(filename):
    is_pdf = filename.lower().endswith('.pdf')
    as_attachment = request.args.get("download") == "1"
    mimetype = 'application/pdf' if is_pdf else None
    response = send_from_directory(
        OUTPUTS_DIR,
        filename,
        mimetype=mimetype,
        as_attachment=as_attachment,
        download_name=filename,
    )
    if is_pdf and not as_attachment:
        response.headers["Content-Disposition"] = f'inline; filename="{filename}"'
    return response


# ── Arranque ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("Starting Flask web app")
    logger.info(f"Template directory: {TEMPLATES_DIR}")
    logger.info(f"Outputs directory: {OUTPUTS_DIR}")
    logger.info(f"Uploads directory: {UPLOADS_DIR}")
    logger.info("=" * 60)
    
    app.run(debug=True, port=5000)
