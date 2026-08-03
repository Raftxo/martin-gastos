"""
app.py — Servidor Flask para la web app de gastos de tacógrafo.

DESCRIPCIÓN GENERAL
-------------------
Esta aplicación web permite cargar un CSV de actividades de tacógrafo,
mapear regiones desconocidas a ciudades, añadir actividades manuales
(opcionalmente) y generar un informe de gastos en Excel/PDF.

FLUJO TÍPICO DE USO
-------------------
1. El usuario abre http://localhost:5000 en el navegador.
2. Selecciona un archivo CSV de actividades y una plantilla Excel.
3. La app analiza el CSV y detecta regiones desconocidas.
4. El usuario completa el mapeo de regiones → ciudades.
5. La app genera el Excel, convierte a PDF y muestra una vista previa.

RUTAS PRINCIPALES
-----------------
- GET  /              : Página principal (index.html).
- GET  /api/templates : Lista las plantillas Excel disponibles.
- POST /api/analyze   : Recibe el CSV y devuelve regiones desconocidas.
- POST /api/generate  : Recibe el formulario completo y genera Excel/PDF.
- GET  /outputs/<f>   : Sirve los archivos generados.

CÓMO EJECUTAR
-------------
    cd web_app
    python app.py

El servidor arranca en http://localhost:5000 y, al iniciarse, abre
automáticamente el navegador en esa dirección.

DEPENDENCIAS
------------
- Flask
- fill_excel_v9 (módulo local)
- Plantillas Excel en ./excel_templates
"""

import os
import json
import logging
import threading
import webbrowser
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from fill_excel_v9 import fill_excel, parse_csv_for_unknowns, convert_to_pdf, export_preview_png, get_destinos_for_province
from parse_tacografo import parse_csv_shifts, extract_destinations

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

TRUCKS_FILE = os.path.join(BASE_DIR, "trucks.json")


def _load_trucks() -> dict:
    """Carga el mapeo matrícula → número interno desde trucks.json."""
    try:
        with open(TRUCKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError as e:
        logger.error(f"Error leyendo {TRUCKS_FILE}: {e}")
        return {}


def _save_trucks(trucks: dict) -> None:
    """Guarda el mapeo matrícula → número interno en trucks.json."""
    try:
        with open(TRUCKS_FILE, "w", encoding="utf-8") as f:
            json.dump(trucks, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logger.error(f"Error guardando {TRUCKS_FILE}: {e}")


def _unknown_trucks(csv_path: str) -> list:
    """Devuelve las matrículas del CSV que no están en trucks.json."""
    trucks = _load_trucks()
    shifts = parse_csv_shifts(csv_path)
    unknowns = set()
    for s in shifts:
        plate = s.get("plate")
        if plate and plate not in trucks:
            unknowns.add(plate)
    return sorted(unknowns)


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


@app.route("/api/destinos/<provincia>")
def get_destinos(provincia):
    """Devuelve los destinos frecuentes para una provincia."""
    destinos = get_destinos_for_province(provincia)
    return jsonify(destinos)


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
        unknown_trucks = _unknown_trucks(csv_path)
        unknown_destinations = extract_destinations(csv_path)
        logger.info(f"Found {len(unknown_regions)} unknown regions: {unknown_regions}")
        logger.info(f"Found {len(unknown_trucks)} unknown trucks: {unknown_trucks}")
        logger.info(f"Found {len(unknown_destinations)} destinations: {unknown_destinations}")

        return jsonify({
            "csv_path": csv_path,
            "unknown_regions": list(unknown_regions),
            "unknown_trucks": unknown_trucks,
            "destinations": unknown_destinations
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
    csv_path            = data.get("csv_path")
    template_name       = data.get("template")
    location_map        = data.get("location_map", {})
    manual_shifts       = data.get("manual_shifts", [])
    truck_map_extra     = data.get("truck_map", {})
    destination_save_map = data.get("destination_save_map", {})  # {"Madrid": "Pinto", "Murcia": "Murcia", ...}

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

    # Guardar en trucks.json los nuevos mapeos de matrículas recibidos
    if truck_map_extra:
        trucks = _load_trucks()
        trucks.update(truck_map_extra)
        _save_trucks(trucks)
        logger.info(f"Saved {len(truck_map_extra)} new truck mappings to {TRUCKS_FILE}")

    try:
        # Guardar destinos frecuentes si se proporcionó un mapa
        if destination_save_map:
            from fill_excel_v9 import save_destino
            for provincia, destino in destination_save_map.items():
                save_destino(provincia, destino)
            logger.info(f"Saved {len(destination_save_map)} destinations to destinos.json")

        output_filename = fill_excel(
            excel_path          = template_path,
            csv_path            = csv_path,
            location_map        = location_map,
            manual_shifts       = manual_shifts,
            output_dir          = OUTPUTS_DIR,
            truck_map_extra     = truck_map_extra,
            destination_save_map = destination_save_map or None,
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
        
        # NOTE: El CSV subido se conserva en uploads/ para facilitar la
        # depuración. Si en el futuro se quiere borrar automáticamente,
        # descomenta el bloque siguiente.
        # try:
        #     if os.path.exists(csv_path):
        #         os.remove(csv_path)
        #         logger.info(f"Cleaned up uploaded CSV: {csv_path}")
        # except Exception as cleanup_err:
        #     logger.warning(f"Failed to cleanup CSV file: {cleanup_err}")

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

def _open_browser_delayed(url: str = "http://localhost:5000", delay: float = 1.0) -> None:
    """Abre el navegador en la URL del servidor tras un pequeño retraso."""
    def _open() -> None:
        logger.info(f"Abriendo navegador en {url}")
        webbrowser.open(url)

    threading.Timer(delay, _open).start()


if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("Starting Flask web app")
    logger.info(f"Template directory: {TEMPLATES_DIR}")
    logger.info(f"Outputs directory: {OUTPUTS_DIR}")
    logger.info(f"Uploads directory: {UPLOADS_DIR}")
    logger.info("=" * 60)

    # Abrir el navegador automáticamente, pero solo en el proceso principal.
    # En modo debug Werkzeug lanza un segundo proceso de recarga; con la
    # variable WERKZEUG_RUN_MAIN evitamos que se abra una segunda ventana.
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        _open_browser_delayed()

    app.run(debug=True, port=5000)
