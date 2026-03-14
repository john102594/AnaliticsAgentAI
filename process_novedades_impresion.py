import os
import argparse
import json
import shutil
import re
from datetime import datetime
import pandas as pd
from prisma import Prisma

# Setup file paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')

def load_config():
    try:
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {CONFIG_FILE}")
        return None

def extract_ot_details(ot_string):
    """Extrae OT, Referencia y Estado de la cadena de OT."""
    # Ejemplo: "OT : 2603152 - Referencia : ME03329 MET 20-17.5 28,5X23 GOLPE MAYO SIN SEL 42G - Estado: Parcial"
    try:
        parts = ot_string.split(" - ")
        ot = parts[0].replace("OT : ", "").strip()
        referencia = parts[1].replace("Referencia : ", "").strip()
        estado = parts[2].replace("Estado: ", "").strip()
        return ot, referencia, estado
    except:
        return "", "", ""

def extract_report_info(report_string):
    """Extrae numero de reporte, fecha y turno"""
    # Ejemplo: Reporte Nro. 60289 - 10/03/2026 9:14:17 - Turno: 1
    try:
        parts = report_string.split(" - ")
        reporte_nro = parts[0].replace("Reporte Nro. ", "").strip()
        fecha_hora_str = parts[1].strip()
        fecha = datetime.strptime(fecha_hora_str, "%d/%m/%Y %H:%M:%S").strftime("%Y-%m-%d")
        turno = int(parts[2].replace("Turno: ", "").strip())
        return str(reporte_nro), fecha, turno
    except:
        return None, None, None

def parse_novedades_file(file_path, sheet_name="Resumen Operarios Imp."):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Analizando archivo: {file_path}")
    try:
        # User validation: ONLY use the configured sheet
        df = pd.read_excel(file_path, engine="xlrd", sheet_name=sheet_name, header=None)
    except Exception as e:
        print(f"Error al leer el archivo Excel {file_path}: {e}")
        return []

    # Clean completely empty rows
    df = df.dropna(how='all')
    
    records = []
    
    current_maquina = None
    current_operario = None
    current_reporte_nro = None
    current_fecha = None
    current_turno = None
    current_obs_inicial = None
    current_obs_final = None
    
    current_ot = None
    current_referencia = None
    current_estado = None
    
    # Structures for the current OT
    variables_t = {}
    detalle_t5 = []
    metros_observacion = None
    desperdicios = []
    
    state = "SEEKING_MACHINE" # SEEKING_MACHINE, IN_MACHINE, IN_OT, IN_DETALLE_T5, IN_METROS, IN_DESPERDICIOS
    
    # Lists for machine's OTs to post-process obs_inicial / obs_final
    machine_ots = []

    def save_current_ot():
        nonlocal machine_ots
        if current_ot and current_maquina and current_fecha:
            # Build the concatenated observaciones string
            # Excluding desperdicios observations as requested
            obs_parts = []
            for t_var, obs in variables_t.items():
                if obs:
                    obs_parts.append(f"Variable {t_var}: {obs}")
            
            for dt5 in detalle_t5:
                obs_parts.append(f"DETALLE T5: Causal {dt5.get('causal','')} Tiempo {dt5.get('tiempo','')} Area {dt5.get('area','')} Obs {dt5.get('observacion','')}")
            
            if metros_observacion:
                obs_parts.append(f"METROS: {metros_observacion}")
            
            observaciones_concat = ", ".join(obs_parts)
            
            machine_ots.append({
                "fecha": current_fecha,
                "turno": current_turno,
                "maquina": current_maquina,
                "operario": current_operario,
                "reporte_nro": current_reporte_nro,
                "obs_inicial": current_obs_inicial,
                "obs_final": current_obs_final,
                "ot": current_ot,
                "referencia": current_referencia,
                "estado_ot": current_estado,
                "observaciones": observaciones_concat,
                "variables_t": variables_t,
                "detalle_t5": detalle_t5,
                "desperdicios": desperdicios
            })
            
    def flush_machine_ots():
        nonlocal machine_ots
        if not machine_ots: return
        
        # Attach obs_inicial to the first OT, obs_final to the last OT
        # Clear them from the others
        obs_i = machine_ots[0]["obs_inicial"]
        obs_f = machine_ots[0]["obs_final"]  # obs_final was captured at the start of the machine block
        
        for idx, ot_rec in enumerate(machine_ots):
            if idx == 0:
                ot_rec["obs_inicial"] = obs_i
            else:
                ot_rec["obs_inicial"] = None
                
            if idx == len(machine_ots) - 1:
                ot_rec["obs_final"] = obs_f
            else:
                ot_rec["obs_final"] = None
                
            records.append(ot_rec)
        
        machine_ots = []
            
    for index, row in df.iterrows():
        col0 = str(row.get(0, "")).strip() if pd.notna(row.get(0)) else ""
        col1 = str(row.get(1, "")).strip() if pd.notna(row.get(1)) else ""
        col2 = str(row.get(2, "")).strip() if pd.notna(row.get(2)) else ""
        col3 = str(row.get(3, "")).strip() if pd.notna(row.get(3)) else ""
        
        if col0.startswith("Máquina :"):
            # If we were in an OT, save it
            save_current_ot()
            flush_machine_ots()
            
            current_maquina = col0.replace("Máquina :", "").strip()
            current_operario = None
            current_reporte_nro = None
            current_fecha = None
            current_turno = None
            current_obs_inicial = None
            current_obs_final = None
            
            current_ot = None
            current_referencia = None
            current_estado = None
            
            variables_t = {}
            detalle_t5 = []
            metros_observacion = None
            desperdicios = []
            
            state = "IN_MACHINE"
            continue
            
        if col0.startswith("Operario :"):
            current_operario = col0.replace("Operario :", "").strip()
            if not current_operario and col1:
                current_operario = col1  # "VARGAS FIGUEROA DARWIN ANDRES" or "Sin registro de Reporte de Novedades"
            if current_operario == "Sin registro de Reporte de Novedades":
                state = "SEEKING_MACHINE"
            continue
            
        if col0.startswith("Reporte Nro."):
            current_reporte_nro, current_fecha, current_turno = extract_report_info(col0)
            continue
            
        if col0.startswith("Observación Inicial del Turno:"):
            current_obs_inicial = col1
            continue
            
        if col0.startswith("Observación Final del Turno:"):
            current_obs_final = col1
            continue
            
        if col0.startswith("OT :"):
            # Save previous OT if exists
            save_current_ot()
            
            # Reset OT variables
            variables_t = {}
            detalle_t5 = []
            metros_observacion = None
            desperdicios = []
            
            current_ot, current_referencia, current_estado = extract_ot_details(col0)
            state = "IN_OT"
            continue
        
        if state == "IN_OT":
            if col0 in ["T1", "T2", "T3", "T4", "T5"]:
                variables_t[col0] = col1 if col1 else None
            elif col0 == "Detalle T5...":
                state = "IN_DETALLE_T5"
            elif col0 == "Metros":
                state = "IN_METROS"
            elif col0 == "Detalle de Desperdicios...":
                state = "IN_DESPERDICIOS"
        
        elif state == "IN_DETALLE_T5":
            if col0 in ["Metros", "Detalle de Desperdicios...", "OT :"] or col0.startswith("Máquina :"):
                if col0 == "Metros": state = "IN_METROS"
                elif col0 == "Detalle de Desperdicios...": state = "IN_DESPERDICIOS"
                elif col0.startswith("OT :"): 
                    save_current_ot()
                    current_ot, current_referencia, current_estado = extract_ot_details(col0)
                    variables_t, detalle_t5, metros_observacion, desperdicios = {}, [], None, []
                    state = "IN_OT"
                elif col0.startswith("Máquina :"):
                    save_current_ot()
                    current_maquina = col0.replace("Máquina :", "").strip()
                    current_operario, current_reporte_nro, current_fecha, current_turno = None, None, None, None
                    current_ot, current_referencia, current_estado = None, None, None
                    variables_t, detalle_t5, metros_observacion, desperdicios = {}, [], None, []
                    state = "IN_MACHINE"
            elif col0 and col0 != "Causal" and col0 != "NaN":
                try:
                    tiempo_val = float(col1) if col1 else 0.0
                except ValueError:
                    tiempo_val = 0.0
                detalle_t5.append({
                    "causal": col0,
                    "tiempo": tiempo_val,
                    "area": col2,
                    "observacion": col3
                })

        elif state == "IN_METROS":
            if col0 == "Observación":
                # The actual observation is usually on the next row in col0
                pass
            elif col0 == "Detalle de Desperdicios...":
                state = "IN_DESPERDICIOS"
            elif col0 and col0 != "NaN":
                # Assuming this is the observation text
                metros_observacion = col0

        elif state == "IN_DESPERDICIOS":
            if col0.startswith("OT :") or col0.startswith("Máquina :") or col0 == "Total Desperdicios:":
                if col0 == "Total Desperdicios:":
                    state = "IN_OT" # done with desperdicios for this OT
                elif col0.startswith("OT :"):
                    save_current_ot()
                    current_ot, current_referencia, current_estado = extract_ot_details(col0)
                    variables_t, detalle_t5, metros_observacion, desperdicios = {}, [], None, []
                    state = "IN_OT"
                elif col0.startswith("Máquina :"):
                    save_current_ot()
                    current_maquina = col0.replace("Máquina :", "").strip()
                    current_operario, current_reporte_nro, current_fecha, current_turno = None, None, None, None
                    current_ot, current_referencia, current_estado = None, None, None
                    variables_t, detalle_t5, metros_observacion, desperdicios = {}, [], None, []
                    state = "IN_MACHINE"
            elif col0 and col0 != "Causal" and col0 != "NaN":
                try:
                    kilos_val = float(col1) if col1 else 0.0
                except ValueError:
                    kilos_val = 0.0
                try:
                    metros_val = float(col2) if col2 else 0.0
                except ValueError:
                    metros_val = 0.0
                    
                desperdicios.append({
                    "causal": col0,
                    "kilos": kilos_val,
                    "metros": metros_val,
                    "observacion": col3
                })

    # Save the last OT if exists
    save_current_ot()
    flush_machine_ots()
    
    return records

def process_and_save_files(file_paths):
    db = Prisma()
    db.connect()
    print("Conectado a la base de datos.")
    
    config = load_config()
    sheet_name = "Resumen Operarios Imp."
    if config:
        reports_cfg = config.get("server", {}).get("reports", [])
        nov_cfg = next((r for r in reports_cfg if r.get("id") == "novedades"), None)
        if nov_cfg and nov_cfg.get("sheet"):
            sheet_name = nov_cfg.get("sheet")
    
    total_processed = 0
    total_inserted = 0
    
    for file_path in file_paths:
        file_name = os.path.basename(file_path)
        print(f"--- Procesando archivo: {file_name} ---")
        try:
            records = parse_novedades_file(file_path, sheet_name=sheet_name)
            total_processed += len(records)
            
            inserted_count = 0
            for record in records:
                # Check uniqueness (fecha, turno, maquina, ot)
                fecha_db = record['fecha']
                
                existing = db.novedadimpresionot.find_first(
                    where={
                        'fecha': fecha_db,
                        'turno': record['turno'],
                        'maquina': record['maquina'],
                        'ot': record['ot']
                    }
                )
                
                if existing:
                    print(f"  Registro ya existe para {record['maquina']}, Turno {record['turno']}, OT {record['ot']}. Ignorando.")
                else:
                    db.novedadimpresionot.create(
                        data={
                            'fecha': fecha_db,
                            'turno': record['turno'],
                            'maquina': record['maquina'],
                            'operario': record['operario'],
                            'reporte_nro': record['reporte_nro'],
                            'obs_inicial': record['obs_inicial'],
                            'obs_final': record['obs_final'],
                            'ot': record['ot'],
                            'referencia': record['referencia'],
                            'estado_ot': record['estado_ot'],
                            'observaciones': record['observaciones'],
                            'variables_t': json.dumps(record['variables_t']),
                            'detalle_ts': json.dumps(record['detalle_t5']),
                            'desperdicios': json.dumps(record['desperdicios']),
                        }
                    )
                    inserted_count += 1
            
            total_inserted += inserted_count
            print(f"  Se encontraron {len(records)} registros en el archivo. Se insertaron {inserted_count} nuevos.")
            
            # Mover archivo procesado a la carpeta _procesados
            processed_dir = os.path.dirname(file_path) + "/_procesados"
            if not os.path.exists(processed_dir):
                os.makedirs(processed_dir)
            
            new_path = os.path.join(processed_dir, file_name)
            shutil.move(file_path, new_path)
            print(f"  Archivo movido a: {new_path}")
            
        except Exception as e:
            print(f"Error procesando el archivo {file_name}: {e}")
            import traceback
            traceback.print_exc()
            
    db.disconnect()
    print("Desconectado de la base de datos.")
    
    return total_processed, total_inserted

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Procesar reportes de novedades de impresión y almacenar en base de datos SQLite.")
    parser.add_argument("--archivo", type=str, help="Ruta a un archivo xls específico para procesar.")
    parser.add_argument("--directorio", type=str, help="Directorio con archivos xls para procesar (por defecto: descargas/novedades/).")
    
    args = parser.parse_args()
    
    config = load_config()
    if not config:
        sys.exit(1)
        
    BASE_DOWNLOAD_DIR = os.path.join(BASE_DIR, config.get("download_dir", "descargas"))
    NOVEDADES_DIR = os.path.join(BASE_DOWNLOAD_DIR, "novedades")
    
    files_to_process = []
    
    if args.archivo:
        if os.path.exists(args.archivo):
            files_to_process.append(args.archivo)
        else:
            print(f"El archivo especificado no existe: {args.archivo}")
            exit(1)
    else:
        target_dir = args.directorio if args.directorio else NOVEDADES_DIR
        if os.path.exists(target_dir):
            for filename in os.listdir(target_dir):
                if filename.endswith(".xls") or filename.endswith(".xlsx"):
                    file_path = os.path.join(target_dir, filename)
                    if os.path.isfile(file_path):
                        files_to_process.append(file_path)
            
            if not files_to_process:
                print(f"No se encontraron archivos .xls en el directorio: {target_dir}")
        else:
            print(f"El directorio especificado no existe: {target_dir}")
            exit(1)
            
    if files_to_process:
        process_and_save_files(files_to_process)
    else:
        print("No hay archivos para procesar.")
