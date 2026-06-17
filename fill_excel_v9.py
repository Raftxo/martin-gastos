"""
fill_excel_v9.py
Versión adaptada para la web app Flask.
Cambios respecto a v8:
  - Eliminados todos los input() interactivos.
  - Nueva función parse_csv_for_unknowns() para el análisis previo.
  - fill_excel() acepta location_map, manual_shifts y output_dir como parámetros.
"""

import os
import pythoncom
from datetime import datetime, time, timedelta
from win32com import client
from parse_tacografo import parse_csv_shifts

# ── Constantes ────────────────────────────────────────────────────────────────

MONTHS = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

TRUCK_MAP = {
    "4974GCV": "T-3204",
    "9529MKG": "T-397",
    "2638KXW": "T-401",
    "5602JWF": "T-403",
    "2084MCH": "T-404",
    "2383KXW": "T-360",
    "2393LHH": "T-144",
    "7028MHL": "T-405",
    "7299NKR": "T423",
    "7599LKN": "T-3206"
}

DIAS_ES = {
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
    "Thursday": "Jueves", "Friday": "Viernes",
    "Saturday": "Sábado", "Sunday": "Domingo"
}

REGION_ABBR = {
    "Andalucía": "(AN)", "Aragón": "(AR)", "Asturias": "(AST)", "Cantabria": "(C)",
    "Cataluña": "(CAT)", "Castilla-León": "(CL)", "Castilla y León": "(CL)",
    "Castilla-La Mancha": "(CM)", "Castilla La Mancha": "(CM)", "Valencia": "(CV)",
    "Comunidad Valenciana": "(CV)", "Extremadura": "(EXT)", "Galicia": "(G)",
    "Baleares": "(IB)", "Islas Baleares": "(IB)", "Canarias": "(IC)", "Islas Canarias": "(IC)",
    "La Rioja": "(LR)", "Madrid": "(M)", "Comunidad de Madrid": "(M)", "Murcia": "(MU)",
    "Región de Murcia": "(MU)", "Navarra": "(NA)", "País Vasco": "(PV)", "Euskadi": "(PV)"
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def write_safe(ws, row, col, value):
    try:
        cell = ws.Cells(row, col)
        if not cell.MergeCells or cell.Address == cell.MergeArea.Item(1).Address:
            cell.Value = value
    except Exception as e:
        print(f"Error escribiendo celda {row},{col}: {e}")

def set_digits(ws, row, start_col, value, length=4):
    s = str(value).replace(':', '').zfill(length)
    for i, digit in enumerate(s):
        write_safe(ws, row, start_col + i, digit)

def set_digits_right_aligned(ws, row, end_col, value):
    s = str(value)
    for i, digit in enumerate(reversed(s)):
        write_safe(ws, row, end_col - i, digit)

# ── Nueva función: análisis previo del CSV ────────────────────────────────────

def parse_csv_for_unknowns(csv_path: str) -> set:
    """
    Lee el CSV y devuelve el conjunto de regiones que no están
    en REGION_ABBR ni son Madrid (que se resuelve automáticamente a 'Pinto').
    El frontend las mostrará al usuario para que introduzca la ciudad.
    """
    shifts = parse_csv_shifts(csv_path)
    unknown = set()
    for s in shifts:
        for loc in [s.get("origin"), s.get("destination")]:
            if not loc:
                continue
            if loc in REGION_ABBR:
                # Madrid se resuelve sola, el resto necesitan ciudad
                if REGION_ABBR[loc] != "(M)":
                    unknown.add(loc)
            else:
                # Región desconocida completamente
                unknown.add(loc)
    return unknown

# ── Función principal ─────────────────────────────────────────────────────────

def fill_excel(
    excel_path: str,
    csv_path: str,
    location_map: dict,      # {"Andalucía": "Sevilla", "Valencia": "Castellón", ...}
    manual_shifts: list,     # [{"fecha": "01/06/2025", "h_ini": "08:00", ...}, ...]
    output_dir: str = ".",
) -> str:
    """
    Genera el Excel de gastos. Devuelve el nombre del fichero generado.
    Ya no usa input() — toda la info viene como parámetros.
    """
    print(f"--> Leyendo CSV desde: {os.path.abspath(csv_path)}")
    shifts = parse_csv_shifts(csv_path)

    # Añadir actividades manuales recibidas desde el formulario web
    for m in manual_shifts:
        try:
            dt_base = datetime.strptime(m["fecha"], "%d/%m/%Y")
            h_i, min_i = map(int, m["h_ini"].split(":"))
            h_f, min_f = map(int, m["h_fin"].split(":"))
            start_dt = dt_base.replace(hour=h_i, minute=min_i)
            end_dt   = dt_base.replace(hour=h_f, minute=min_f)
            shifts.append({
                "start_dt":      start_dt,
                "end_dt":        end_dt,
                "plate":         "COCHE PARTICULAR",
                "km_start":      None,
                "km_end":        None,
                "km_total":      m.get("km", ""),
                "origin":        m.get("concepto", ""),
                "destination":   m.get("concepto", ""),
                "work_duration": end_dt - start_dt,
                "drive_duration": timedelta(0),
            })
        except Exception as e:
            print(f"[WARN] Actividad manual ignorada por error: {e}")

    shifts.sort(key=lambda x: x["start_dt"])

    if not shifts:
        raise ValueError("No se encontraron jornadas en el CSV.")

    # Pre-calcular inicio/fin total de cada día
    daily_stats = {}
    for s in shifts:
        d = s["start_dt"].date()
        s_time = s["start_dt"]
        e_time = s["end_dt"] if s["end_dt"] else s["start_dt"]
        if d not in daily_stats:
            daily_stats[d] = {"start": s_time, "end": e_time}
        else:
            if s_time < daily_stats[d]["start"]: daily_stats[d]["start"] = s_time
            if e_time > daily_stats[d]["end"]:   daily_stats[d]["end"] = e_time

    # Resolver ubicaciones usando el mapa recibido + reglas fijas
    def resolve_location(raw_loc):
        if not raw_loc:
            raw_loc = "Madrid"
        abbr = REGION_ABBR.get(raw_loc, "")
        if abbr == "(M)":
            return "Pinto"
        # Si el usuario proporcionó una ciudad para esta región, la usamos
        if raw_loc in location_map:
            city = location_map[raw_loc]
            suffix = abbr if abbr else ""
            return f"{city} {suffix}".strip()
        # Si ya tiene abreviatura pero no se proporcionó ciudad, devolvemos la región
        if abbr:
            return f"{raw_loc} {abbr}"
        return raw_loc

    # ── COM Excel ────────────────────────────────────────────────────────────
    T_0700 = time(7, 0)
    T_1300 = time(13, 0)
    T_1500 = time(15, 0)
    T_2200 = time(22, 0)

    assigned_meals = {}
    wb_template = None
    wb_output   = None

    try:
        pythoncom.CoInitialize()
        excel = client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_template = os.path.abspath(excel_path)
        wb_template   = excel.Workbooks.Open(abs_template)
        template_sheet = wb_template.Sheets(1)

        wb_output = excel.Workbooks.Add()

        ws = None
        dia_row       = 6
        step          = 7
        jornada_index = 1

        for shift in shifts:
            mes_nombre = MONTHS[shift["start_dt"].month].upper()

            if jornada_index == 1 or (jornada_index - 1) % 7 == 0:
                template_sheet.Copy(Before=wb_output.Sheets(1))
                ws = wb_output.Sheets(1)
                numero_hoja = ((jornada_index - 1) // 7) + 1
                ws.Name = f"{mes_nombre} ({numero_hoja})"
                dia_row = 6
                write_safe(ws, 1, 35, mes_nombre)
                write_safe(ws, 1, 39, shift["start_dt"].year)

            weekday_es = DIAS_ES[shift["start_dt"].strftime("%A")]
            write_safe(ws, dia_row - 2, 1, weekday_es)
            write_safe(ws, dia_row,     1, shift["start_dt"].day)
            write_safe(ws, dia_row,     3, shift["start_dt"].month)
            write_safe(ws, dia_row + 2, 2, jornada_index)

            set_digits(ws, dia_row - 2, 6, shift["start_dt"].strftime("%H:%M"))
            if shift["end_dt"]:
                set_digits(ws, dia_row - 1, 6, shift["end_dt"].strftime("%H:%M"))

            write_safe(ws, dia_row - 2, 11, resolve_location(shift["origin"]))
            write_safe(ws, dia_row - 1, 11, resolve_location(shift["destination"] or shift["origin"]))

            current_date = shift["start_dt"].date()
            day_start    = daily_stats[current_date]["start"]
            day_end      = daily_stats[current_date]["end"]

            plate = shift["plate"]
            plate_text = f"{plate} / {TRUCK_MAP[plate]}" if plate in TRUCK_MAP else plate

            if (day_end - day_start) >= timedelta(hours=12):
                plate_text += "\ndietas 12 horas"
                ws.Cells(dia_row + 1, 5).WrapText = True  # type: ignore[union-attr]

            write_safe(ws, dia_row + 1, 5, plate_text)

            # KMs
            if shift.get("plate") == "COCHE PARTICULAR":
                km_total = shift.get("km_total")
                if km_total:
                    set_digits_right_aligned(ws, dia_row + 3, 20, km_total)
            else:
                km_start = shift["km_start"]
                km_end   = shift["km_end"]
                if km_end:   set_digits_right_aligned(ws, dia_row + 1, 20, km_end)
                if km_start: set_digits_right_aligned(ws, dia_row + 2, 20, km_start)
                if km_start and km_end:
                    set_digits_right_aligned(ws, dia_row + 3, 20, km_end - km_start)

            # Dietas
            if current_date not in assigned_meals:
                assigned_meals[current_date] = set()

            s_start_t = shift["start_dt"].time()
            s_end_t   = (shift["end_dt"] if shift["end_dt"] else shift["start_dt"]).time()
            s_end_dt  = shift["end_dt"]  if shift["end_dt"]  else shift["start_dt"]

            shift_meals = set()
            if s_start_t < T_0700:                        shift_meals.add("desayuno")
            if s_start_t < T_1500 and s_end_t > T_1300:  shift_meals.add("comida")
            if s_end_t > T_2200:                          shift_meals.add("cena")

            if (day_end - day_start) >= timedelta(hours=12):
                if s_end_dt == day_end:
                    for meal in ("desayuno", "comida", "cena"):
                        if meal not in assigned_meals[current_date]:
                            shift_meals.add(meal)
                assigned_meals[current_date].add("note_12h")

            if "desayuno" in shift_meals and "desayuno" not in assigned_meals[current_date]:
                write_safe(ws, dia_row + 1, 24, "X")
                assigned_meals[current_date].add("desayuno")
            if "comida" in shift_meals and "comida" not in assigned_meals[current_date]:
                write_safe(ws, dia_row + 1, 25, "X")
                assigned_meals[current_date].add("comida")
            if "cena" in shift_meals and "cena" not in assigned_meals[current_date]:
                write_safe(ws, dia_row + 1, 26, "X")
                assigned_meals[current_date].add("cena")

            dia_row += step
            jornada_index += 1

        # Borrar hoja en blanco por defecto
        if wb_output.Sheets.Count > 1:
            wb_output.Sheets(wb_output.Sheets.Count).Delete()

        # Guardar
        first_date     = shifts[0]["start_dt"]
        year           = first_date.year
        month_name     = MONTHS[first_date.month].lower()[:3]
        num_hojas      = wb_output.Sheets.Count
        output_filename = f"Gastos_0529_RJW_{year}_{month_name}({num_hojas}hojas).xlsx"

        abs_output = os.path.join(os.path.abspath(output_dir), output_filename)
        wb_output.SaveAs(abs_output)
        print(f"Generado: {output_filename}")
        return output_filename

    except Exception as e:
        raise
    finally:
        if wb_template: wb_template.Close(False)
        if wb_output:   wb_output.Close(False)
        if 'excel' in dir():
            excel.Quit()
        pythoncom.CoUninitialize()

# ── Función de conversión a PDF ─────────────────────────────────────────────────────────

def convert_to_pdf(excel_file_path):
    pythoncom.CoInitialize()
    abs_path = os.path.abspath(excel_file_path)
    pdf_path = abs_path.replace('.xlsx', '.pdf')
    
    excel = client.Dispatch("Excel.Application")
    excel.Visible = False
    
    try:
        wb_com = excel.Workbooks.Open(abs_path)
        wb_com.ExportAsFixedFormat(0, pdf_path)
        print(f"Generado PDF: {os.path.basename(pdf_path)}")
    finally:
        if 'wb_com' in locals():
            wb_com.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()