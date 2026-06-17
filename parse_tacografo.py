import csv
import re
from datetime import datetime, timedelta

def parse_duration(duration_str):
    if not duration_str or duration_str == "":
        return timedelta(0)
    h, m = map(int, duration_str.split(':'))
    return timedelta(hours=h, minutes=m)

def parse_csv_shifts(file_path):
    shifts = []
    
    with open(file_path, mode='r', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter=';')
        
        current_shift = None
        
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
            
            # Location events (can be origin or destination)
            elif "#SYS#COUNTRY" in description:
                loc_match = re.search(r'\|,\s*(.*)', description)
                if loc_match:
                    location = loc_match.group(1).split('(')[0].strip()
                    km_match = re.search(r'\((\d+)km\)', description)
                    km = int(km_match.group(1)) if km_match else None
                    
                    if current_shift:
                        if current_shift['origin'] is None:
                            current_shift['origin'] = location
                        else:
                            current_shift['destination'] = location
                        
                        if km and (current_shift['km_end'] is None or km > current_shift['km_end']):
                            current_shift['km_end'] = km
                            current_shift['end_dt'] = dt # If it has KM, it might be the end of day

            # Card Out (Explicit end)
            elif "Card Out" in description:
                match = re.search(r'Card Out \((.*?), (\d+)km\)', description)
                if current_shift:
                    current_shift['km_end'] = int(match.group(2)) if match else current_shift['km_end']
                    current_shift['end_dt'] = dt
                    shifts.append(current_shift)
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

if __name__ == "__main__":
    csv_file = "ACTIVITIES_RAFAL JANUSZ_WYSOCKI_EX1434298H000003_20260225_172525.csv"
    results = parse_csv_shifts(csv_file)
    for s in results:
        print(f"Start: {s['start_dt']} | Plate: {s['plate']} | KMs: {s['km_start']} - {s['km_end']} | {s['origin']} -> {s['destination']}")
