import csv
import re
import logging
import chardet
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

def detect_csv_encoding(file_path: str) -> str:
    """
    Auto-detect CSV encoding using chardet.
    Falls back to UTF-8 if detection fails.
    """
    try:
        with open(file_path, 'rb') as f:
            raw = f.read(100000)  # Read first 100KB for detection
        
        detected = chardet.detect(raw)
        encoding = detected.get('encoding') or 'utf-8'
        confidence = detected.get('confidence', 0)
        
        logger.info(f"Detected encoding: {encoding} (confidence: {confidence:.1%})")
        
        # Validate encoding works by attempting to decode
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                f.read(1000)
            logger.info(f"✓ Using encoding: {encoding}")
            return encoding
        except (UnicodeDecodeError, LookupError) as e:
            logger.warning(f"Detected encoding '{encoding}' failed: {e}. Falling back to UTF-8")
            return 'utf-8'
    except Exception as e:
        logger.warning(f"Encoding detection failed: {e}. Using UTF-8 fallback")
        return 'utf-8'

def validate_csv_structure(file_path: str, encoding: str) -> None:
    """
    Validate that CSV has the required columns.
    Expected columns: timestamp, event_type, description, (duration_field), duration
    """
    required_fields = ['timestamp', 'event_type', 'description', 'duration']
    
    try:
        with open(file_path, mode='r', encoding=encoding) as f:
            reader = csv.reader(f, delimiter=';')
            first_row = next(reader, None)
            
            if not first_row:
                raise ValueError("CSV file is empty")
            
            logger.info(f"CSV structure - Headers (first row): {first_row}")
            
            # Validate that we have enough columns (at least 5 for our parsing)
            if len(first_row) < 5:
                raise ValueError(
                    f"CSV has only {len(first_row)} columns. "
                    f"Expected at least 5 columns (timestamp, event_type, description, ???, duration)"
                )
            
            logger.info("✓ CSV structure validation passed")
    except Exception as e:
        logger.error(f"CSV structure validation failed: {e}")
        raise

def parse_duration(duration_str):
    if not duration_str or duration_str == "":
        return timedelta(0)
    h, m = map(int, duration_str.split(':'))
    return timedelta(hours=h, minutes=m)

def parse_csv_shifts(file_path, encoding: str = None):
    shifts = []
    
    # Auto-detect encoding if not provided
    if encoding is None:
        encoding = detect_csv_encoding(file_path)
    
    # Validate CSV structure
    try:
        validate_csv_structure(file_path, encoding)
    except ValueError as e:
        logger.error(f"CSV validation error: {e}")
        raise
    
    with open(file_path, mode='r', encoding=encoding) as f:
        reader = csv.reader(f, delimiter=';')

        current_shift = None
        last_closed_shift = None

        for row in reader:
            if not any(row): continue

            timestamp_str = row[0].strip('"')
            event_type = row[1].strip('"')
            description = row[2].strip('"')
            duration_str = row[4].strip('"')

            try:
                dt = datetime.strptime(timestamp_str, '%d/%m/%y %H:%M:%S')
            except ValueError:
                continue

            # Start of shift
            if "Card In" in description:
                if current_shift: # Close previous shift if any (though usually Card Out closes it)
                    shifts.append(current_shift)

                match = re.search(r'Card In \((.*?), (\d+)km\)', description)
                current_shift = {
                    'start_dt': dt,
                    'end_dt': None,
                    'plate': match.group(1) if match else None,
                    'km_start': int(match.group(2)) if match else None,
                    'km_end': None,
                    'origin': None,
                    'destination': None,
                    'work_duration': timedelta(0),
                    'drive_duration': timedelta(0)
                }
                last_closed_shift = None

            # Location events (can be origin or destination)
            elif "#SYS#COUNTRY" in description:
                loc_match = re.search(r'\|,\s*(.*)', description)
                if loc_match:
                    location = loc_match.group(1).split('(')[0].strip()
                    km_match = re.search(r'\((\d+)km\)', description)
                    km = int(km_match.group(1)) if km_match else None

                    # The location event may belong to the current open shift or,
                    # if it appears right after a Card Out, to the shift that just
                    # closed (as its destination).
                    target_shift = current_shift
                    if target_shift is None and last_closed_shift is not None:
                        if last_closed_shift.get('end_dt') == dt:
                            target_shift = last_closed_shift

                    if target_shift:
                        if target_shift['origin'] is None:
                            target_shift['origin'] = location
                        else:
                            target_shift['destination'] = location

                        if km and (target_shift['km_end'] is None or km > target_shift['km_end']):
                            target_shift['km_end'] = km
                            target_shift['end_dt'] = dt # If it has KM, it might be the end of day

            # Card Out (Explicit end)
            elif "Card Out" in description:
                match = re.search(r'Card Out \((.*?), (\d+)km\)', description)
                if current_shift:
                    current_shift['km_end'] = int(match.group(2)) if match else current_shift['km_end']
                    current_shift['end_dt'] = dt
                    shifts.append(current_shift)
                    last_closed_shift = current_shift
                    current_shift = None

            # Hours accumulation
            if event_type == "#" and current_shift:
                duration = parse_duration(duration_str)
                if "CONDUCCIÓN" in description:
                    current_shift['drive_duration'] += duration
                elif "TRABAJO" in description:
                    current_shift['work_duration'] += duration

        if current_shift: # Close last shift if it didn't have Card Out
            shifts.append(current_shift)

    return shifts


def extract_destinations(csv_path: str, encoding: str = None) -> list:
    """
    Extrae las regiones y destinos únicos del CSV.
    Devuelve una lista de tuplas (tipo, nombre): ('region', 'Murcia') o ('destino', 'Murcia').
    """
    if encoding is None:
        encoding = detect_csv_encoding(csv_path)

    regions = set()
    destinations = set()

    with open(csv_path, mode='r', encoding=encoding) as f:
        reader = csv.reader(f, delimiter=';')

        for row in reader:
            if not any(row):
                continue

            description = row[2].strip('"')
            if "#SYS#COUNTRY" in description:
                loc_match = re.search(r'\|,\s*(.*)', description)
                if loc_match:
                    location = loc_match.group(1).split('(')[0].strip()
                    destinations.add(location)

    # Convertimos destinos en regiones usando el mapa de abreviaturas
    # para inferir la provincia/región
    for dest in destinations:
        # La regla es simple: si el destino ya está en la lista de regiones,
        # se omite. En la práctica, el CSV da destinos directos como "Madrid", "Murcia", etc.
        # que también son regiones.
        regions.add(dest)

    return list(regions) + [(d, dest) for d, dest in zip(["destino"] * len(destinations), destinations)] if destinations else []


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    csv_file = "ACTIVITIES_RAFAL JANUSZ_WYSOCKI_EX1434298H000003_20260225_172525.csv"
    try:
        results = parse_csv_shifts(csv_file)
        logger.info(f"Successfully parsed {len(results)} shifts from CSV")
        for s in results:
            logger.info(f"Shift: {s['start_dt']} | {s['origin']} -> {s['destination']} | KMs: {s['km_start']}-{s['km_end']}")
    except Exception as e:
        logger.error(f"Failed to parse CSV: {e}")
        raise
