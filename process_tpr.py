import os
import shutil
import json
import argparse
import pandas as pd
import datetime
from prisma import Prisma

def safe_str(val):
    if pd.isna(val) or val is None:
        return None
    return str(val).strip()

def safe_float(val):
    if pd.isna(val) or val is None:
        return None
    try:
        val_str = str(val).replace(',', '.').strip()
        if not val_str:
            return None
        return float(val_str)
    except ValueError:
        return None

def load_config(config_path="config.json"):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(script_dir, config_path)
    try:
        with open(full_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Error cargando config.json: {e}")
        return {}


def process_and_save_files(file_paths=None):
    config = load_config()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    target_dir = os.path.join(script_dir, config.get("descargas_dir", "descargas"), "TPR")
    procesados_dir = os.path.join(target_dir, "_procesados")
    
    if not os.path.exists(procesados_dir):
        os.makedirs(procesados_dir)

    reports_cfg = config.get("server", {}).get("reports", [])
    tpr_cfg = next((r for r in reports_cfg if r.get("id") == "tpr"), None)
    
    sheets_mapping = {
        "Detalle Corte": "corte",
        "Det Sellado": "sellado",
        "Det Laminacion": "laminacion",
        "Det Impresion": "impresion",
        "Detalle EXTLAM": "extlam",
        "Detalle Ext": "extrusion",
        "Detalle Montaje": "montaje"
    }
    
    if tpr_cfg and tpr_cfg.get("sheets_to_process"):
        sheets_mapping = tpr_cfg.get("sheets_to_process")

    if not file_paths:
        file_paths = []
        if os.path.exists(target_dir):
            for filename in os.listdir(target_dir):
                if filename.endswith(".xlsx") or filename.endswith(".xls"):
                    file_path = os.path.join(target_dir, filename)
                    if os.path.isfile(file_path):
                        file_paths.append(file_path)

    if not file_paths:
        print("No hay archivos de TPR para procesar.")
        return

    db = Prisma()
    db.connect()
    print("Conectado a la base de datos.")

    try:
        for file_path in file_paths:
            filename = os.path.basename(file_path)
            print(f"\n--- Procesando archivo: {filename} ---")
            
            # Extraer fecha del nombre del archivo (ej: 20260115)
            fecha_archivo = None
            try:
                parts = filename.replace(".xlsx", "").split("_")
                fecha_archivo = parts[-1] 
            except:
                pass

            # Cargamos registros existentes para evitar duplicados (ID compuesto de columnas)
            # Para optimizar, podríamos filtrar por el mes del archivo, pero para mayor seguridad 
            # buscaremos en los registros recientes.
            existing_records = db.tpr.find_many(
                where={"proceso": {"in": list(sheets_mapping.values())}},
                take=100000 # Un límite razonable para comparación en memoria
            )
            
            # Generamos un set de "firmas" únicas: (maquina, fecha, orden, causal, tiempo)
            existing_hashes = set()
            for r in existing_records:
                key = (r.maquina, r.fecha_reporte, r.orden, r.causal, r.tiempo)
                existing_hashes.add(key)
            
            print(f"  Cargados {len(existing_hashes)} registros existentes para validación de duplicados.")

            print(f"  Abriendo archivo excel {filename}...")
            try:
                # Usar context manager para asegurar que el archivo se cierre y se pueda mover
                with pd.ExcelFile(file_path, engine="openpyxl") as excel_file:
                    print(f"  Archivo excel abierto.")
                    
                    for sheet_name, proceso_name in sheets_mapping.items():
                        if sheet_name not in excel_file.sheet_names:
                            continue
                        
                        try:
                            print(f"  Leyendo hoja '{sheet_name}'...")
                            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=1)
                        except Exception as e:
                            print(f"  Error al leer hoja '{sheet_name}': {e}")
                            continue
                            
                        if df.empty:
                            continue

                        # Mapeo de columnas
                        mapper = {}
                        for col in df.columns:
                            c = str(col).strip()
                            if c == "Maquina": mapper["maquina"] = col
                            elif c == "Fecha": mapper["fecha"] = col
                            elif c == "Horas": mapper["horas"] = col
                            elif c == "Causal": mapper["causal"] = col
                            elif c == "Tiempo" and "tiempo" not in mapper: mapper["tiempo"] = col
                            elif c.startswith("Tiempo."): mapper["tiempo_2"] = col
                            elif c == "Responsable": mapper["responsable"] = col
                            elif c == "Orden": mapper["orden"] = col
                            elif c == "Referencia": mapper["referencia"] = col
                            elif c == "Cliente": mapper["cliente"] = col
                            elif c == "Usuario": mapper["usuario"] = col
                            elif c == "Observaciones": mapper["observaciones"] = col
                            elif c == "Complejidad": mapper["complejidad"] = col
                            elif c == "Ficha": mapper["ficha"] = col
                            elif c == "Edicion": mapper["edicion"] = col
                            elif c == "Cyrel": mapper["cyrel"] = col
                            elif c == "Etapa": mapper["etapa"] = col

                        batch_data = []
                        skipped = 0
                        inserted = 0
                        
                        print(f"  Analizando {len(df)} filas en {sheet_name}...")
                        for _, row in df.iterrows():
                            maquina = safe_str(row.get(mapper.get("maquina", "Maquina")))
                            if not maquina: continue

                            # Formatear Fecha
                            raw_f = row.get(mapper.get("fecha", "Fecha"))
                            fecha_rep = None
                            if pd.notna(raw_f):
                                if isinstance(raw_f, (datetime.date, datetime.datetime, pd.Timestamp)):
                                    fecha_rep = raw_f.strftime('%Y-%m-%d')
                                else:
                                    try: fecha_rep = pd.to_datetime(raw_f).strftime('%Y-%m-%d')
                                    except: fecha_rep = str(raw_f).split(" ")[0]

                            orden = safe_str(row.get(mapper.get("orden", "Orden")))
                            causal = safe_str(row.get(mapper.get("causal", "Causal")))
                            tiempo = safe_float(row.get(mapper.get("tiempo", "Tiempo")))
                            
                            # Verificar si existe
                            row_key = (maquina, fecha_rep, orden, causal, tiempo)
                            if row_key in existing_hashes:
                                skipped += 1
                                continue

                            # Si no existe, preparar para insertar
                            batch_data.append({
                                "proceso": proceso_name,
                                "maquina": maquina,
                                "fecha_reporte": fecha_rep,
                                "horas": safe_float(row.get(mapper.get("horas"))),
                                "causal": causal,
                                "tiempo": tiempo,
                                "responsable": safe_str(row.get(mapper.get("responsable"))),
                                "orden": orden,
                                "referencia": safe_str(row.get(mapper.get("referencia"))),
                                "cliente": safe_str(row.get(mapper.get("cliente"))),
                                "usuario": safe_str(row.get(mapper.get("usuario"))),
                                "observaciones": safe_str(row.get(mapper.get("observaciones"))),
                                "complejidad": safe_str(row.get(mapper.get("complejidad"))),
                                "tiempo_2": safe_float(row.get(mapper.get("tiempo_2"))),
                                "ficha": safe_str(row.get(mapper.get("ficha"))),
                                "edicion": safe_str(row.get(mapper.get("edicion"))),
                                "cyrel": safe_str(row.get(mapper.get("cyrel"))),
                                "etapa": safe_str(row.get(mapper.get("etapa"))),
                                "fecha_archivo": fecha_archivo,
                                "datos": json.dumps({k:v for k,v in row.to_dict().items() if pd.notna(v)}, default=str)
                            })
                            
                            # Mantener el set actualizado para este mismo archivo
                            existing_hashes.add(row_key)

                            if len(batch_data) >= 500:
                                db.tpr.create_many(data=batch_data)
                                inserted += len(batch_data)
                                batch_data = []

                        if batch_data:
                            db.tpr.create_many(data=batch_data)
                            inserted += len(batch_data)

                        print(f"    -> {sheet_name}: {inserted} nuevos, {skipped} duplicados saltados.")

            except Exception as e:
                print(f"  Error procesando {filename}: {e}")
                continue
            
            # Mover archivo (ahora que el 'with' cerró el handle)
            ruta_dest = os.path.join(procesados_dir, filename)
            try:
                shutil.move(file_path, ruta_dest)
                print(f"  Archivo movido a _procesados correctamente.")
            except Exception as e:
                print(f"  Error al mover archivo: {e}")

    finally:
        db.disconnect()
        print("Desconectado.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--archivo", type=str)
    args = parser.parse_args()
    process_and_save_files([args.archivo] if args.archivo else None)
