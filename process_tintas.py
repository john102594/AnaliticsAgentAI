import os
import shutil
import json
import argparse
import pandas as pd
from prisma import Prisma

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
    target_dir = os.path.join(script_dir, config.get("descargas_dir", "descargas"), "tintas")
    procesados_dir = os.path.join(target_dir, "_procesados")
    
    if not os.path.exists(procesados_dir):
        os.makedirs(procesados_dir)
        print(f"Directorio creado: {procesados_dir}")

    # Extraer el sheet_name configurado
    sheet_name = "Acumulado"
    reports_cfg = config.get("server", {}).get("reports", [])
    tintas_cfg = next((r for r in reports_cfg if r.get("id") == "tintas"), None)
    if tintas_cfg and tintas_cfg.get("sheet"):
        sheet_name = tintas_cfg.get("sheet")

    if not file_paths:
        file_paths = []
        if os.path.exists(target_dir):
            for filename in os.listdir(target_dir):
                if filename.endswith(".xlsx") or filename.endswith(".xls"):
                    file_path = os.path.join(target_dir, filename)
                    if os.path.isfile(file_path):
                        file_paths.append(file_path)

    if not file_paths:
        print("No hay archivos de tintas para procesar.")
        return

    db = Prisma()
    db.connect()
    print("Conectado a la base de datos.")

    try:
        for file_path in file_paths:
            filename = os.path.basename(file_path)
            print(f"\n--- Procesando archivo: {filename} ---")
            
            # The user requested to read 'Acumulado' sheet, headers start at row 3 (skip first 2 rows).
            # Which means header=2 in pandas (0-indexed).
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=2, engine="openpyxl")
            except Exception as e:
                print(f"Error al leer hoja '{sheet_name}' en {filename}: {e}")
                continue
                
            if df.empty:
                print(f"  El archivo está vacío.")
                continue

            insertados = 0
            
            for index, row in df.iterrows():
                # Orden is our primary unique key for the report line
                orden_val = row.get("Orden")
                if pd.isna(orden_val) or not str(orden_val).strip():
                    continue

                orden = str(int(orden_val)) if isinstance(orden_val, float) else str(orden_val).strip()

                def s_float(col_name):
                    val = row.get(col_name)
                    if pd.isna(val) or val is None:
                        return None
                    try:
                        return float(val)
                    except ValueError:
                        return None
                        
                def s_str(col_name):
                    val = row.get(col_name)
                    if pd.isna(val) or val is None:
                        return None
                    return str(val).strip()

                nuevo_registro = {
                    "orden": orden,
                    "referencia": s_str("Referencia"),
                    "linea_financiera": s_str("Linea Financiera"),
                    "kilos_tinta": s_float("Kilos Tinta"),
                    "valor_tintas": s_float("Valor Tintas"),
                    "real_tintas_pct": s_float("% Real Tintas"),
                    "kilos_impresos": s_float("Kilos Impresos"),
                    "metros_impresos": s_float("Metros Impresos"),
                    "kls_std_tintas": s_float("Kls. Std. Tintas (o.t)"),
                    "mts_impreso_kg_std": s_float("Mts. impreso / Kg Std Tintas"),
                    "mts_impreso_kg_real": s_float("Mts. impreso / Kg Real Tintas"),
                    "dif_real_tin_vs_tin_ot": s_float("Dif. Real Tin vs Tin o.t."),
                    "dif_consumo_kls_st_pct": s_float("Dif. Consumo/Kls st(%)"),
                    "kilt_real": s_float("KILT_REAL"),
                    "kilt_std": s_float("KILT_STD"),
                    "kils_imp": s_float("KILS_IMP"),
                    "kils_std": s_float("KILS_STD"),
                    "kild_imp": s_float("KILD_IMP"),
                    "kild_std": s_float("KILD_STD"),
                    "mts_imp": s_float("MTS_IMP"),
                    "mts_std": s_float("MTS_STD"),
                    "mtsd_imp": s_float("MTSD_IMP"),
                    "mtsd_std": s_float("MTSD_STD"),
                    "cyrel": s_float("cyrel")
                }

                # Check if it exists
                existente = db.consumotintas.find_unique(where={"orden": orden})

                if not existente:
                    try:
                        db.consumotintas.create(data=nuevo_registro)
                        insertados += 1
                    except Exception as e:
                        print(f"Error insertando la OT {orden}: {e}")
                else:
                    # You can also update if they already exist, but according to user:
                    # "ya abran datos en la bd con los mismos valores esos no se deben procesar"
                    pass

            print(f"  Procesado. Se insertaron {insertados} registros nuevos (Acumulados guardados en tabla).")
            
            # Post-procesar y mover el archivo
            ruta_destino = os.path.join(procesados_dir, filename)
            try:
                shutil.move(file_path, ruta_destino)
            except Exception as e:
                print(f"  No se pudo mover a _procesados: {e}")

    finally:
        db.disconnect()
        print("Desconectado de la base de datos.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Procesa reportes de Consumo de Tintas acumulados descargados y los guarda en base de datos SQLite.")
    parser.add_argument("--archivo", type=str, help="Ruta de un archivo específico (.xlsx) a procesar. Opcional.")
    args = parser.parse_args()

    if args.archivo:
        if os.path.isfile(args.archivo):
            process_and_save_files([args.archivo])
        else:
            print(f"El archivo especificado no existe: {args.archivo}")
    else:
        process_and_save_files()
