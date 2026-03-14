import os
import shutil
import json
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

def process_desperdicios():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Directorios de origen y destino
    descargas_dir = os.path.join(script_dir, "descargas", "desperdicios")
    procesados_dir = os.path.join(descargas_dir, "_procesados")
    
    if not os.path.exists(descargas_dir):
        print(f"El directorio {descargas_dir} no existe. No hay nada que procesar.")
        return
        
    if not os.path.exists(procesados_dir):
        os.makedirs(procesados_dir)
        print(f"Creado directorio para procesados: {procesados_dir}")

    # Load config to get dynamic sheet name
    sheet_name = "Desperdicio Liquidados Turno"
    try:
        config_path = os.path.join(script_dir, "config.json")
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            reports = cfg.get("server", {}).get("reports", [])
            desp_cfg = next((r for r in reports if r.get("id") == "desperdicios"), None)
            if desp_cfg and desp_cfg.get("sheet"):
                sheet_name = desp_cfg.get("sheet")
    except Exception as e:
        print(f"Error cargando config para sheet_name de desperdicios: {e}")

    archivos = [f for f in os.listdir(descargas_dir) if f.endswith(".xlsx") and os.path.isfile(os.path.join(descargas_dir, f))]
    
    if not archivos:
        print("No se encontraron archivos de desperdicios para procesar.")
        return
        
    db = Prisma()
    db.connect()
    
    num_archivos_procesados = 0

    try:
        for archivo in archivos:
            ruta_archivo = os.path.join(descargas_dir, archivo)
            ruta_destino = os.path.join(procesados_dir, archivo)
            print(f"\n--- Procesando archivo: {archivo} ---")
            
            try:
                # La tabla empieza en la celda A5 (Fila 5 en excel) lo que es header=4 en pandas index=(0,1,2,3,4)
                df = pd.read_excel(ruta_archivo, sheet_name=sheet_name, header=4, engine="openpyxl")
                
                # Filtrar filas vacías si las hay
                # Asumo que la columna 'Fecha' o 'Maquina' deben tener datos
                if "Fecha" in df.columns:
                    df = df.dropna(subset=["Fecha"])
                elif "Maquina" in df.columns:
                    df = df.dropna(subset=["Maquina"])
                
                if df.empty:
                    print(f"El archivo {archivo} no contiene datos válidos para procesar.")
                    shutil.move(ruta_archivo, ruta_destino)
                    continue

                registros_creados = 0
                registros_omitidos = 0

                for _, row in df.iterrows():
                    # Formateo de las fechas si es tipo Timestamp (Pandas)
                    raw_fecha = row.get("Fecha")
                    if pd.notna(raw_fecha):
                        if isinstance(raw_fecha, str):
                            fecha = raw_fecha.split(" ")[0] # Puede tener formato YYYY-MM-DD HH:MM:SS
                        elif isinstance(raw_fecha, (datetime.date, datetime.datetime, pd.Timestamp)):
                            fecha = raw_fecha.strftime('%Y-%m-%d')
                        else:
                            try:
                                fecha = pd.to_datetime(raw_fecha).strftime('%Y-%m-%d')
                            except Exception:
                                fecha = str(raw_fecha).split(" ")[0]
                    else:
                        fecha = None
                        
                    raw_hora = row.get("Hora")
                    if pd.notna(raw_hora):
                        if isinstance(raw_hora, str):
                            hora = raw_hora
                        elif isinstance(raw_hora, (datetime.time, datetime.datetime, pd.Timestamp)):
                            hora = raw_hora.strftime('%H:%M:%S')
                        else:
                            try:
                                hora = pd.to_datetime(raw_hora).strftime('%H:%M:%S')
                            except Exception:
                                hora = str(raw_hora)
                    else:
                        hora = None
                    
                    turno_val = safe_float(row.get("Turno"))
                    turno = int(turno_val) if turno_val is not None else None
                    proceso = safe_str(row.get("Proceso"))
                    maquina = safe_str(row.get("Maquina"))
                    bodega = safe_str(row.get("Bodega"))
                    o_trabajo = safe_str(row.get("O.Trabajo"))
                    referencia = safe_str(row.get("Referencia"))
                    item_silp = safe_str(row.get("Item_Silp"))
                    descripcion = safe_str(row.get("Descripcion"))
                    kilos = safe_float(row.get("Kilos"))
                    usuario = safe_str(row.get("Usuario Registra"))
                    observacion = safe_str(row.get("Observacion"))
                    kilos_rech = safe_float(row.get("Kilos_Rech"))
                    rechazado_por = safe_str(row.get("Rechazado Por"))
                    observ_rechz = safe_str(row.get("Observ_Rechz"))
                    estado = safe_str(row.get("Estado"))
                    causal = safe_str(row.get("Causal"))

                    if not fecha: 
                        continue # Sin fecha no insertamos

                    # Validar existencia antes de crear para no duplicar datos
                    # Buscamos un registro idéntico en fecha, hora, turno, maquina, O.Trabajo, Causal y kilos
                    busqueda = db.desperdicio.find_first(
                        where={
                            "fecha": fecha,
                            "hora": hora,
                            "turno": turno,
                            "maquina": maquina,
                            "o_trabajo": o_trabajo,
                            "causal": causal,
                            "kilos": kilos
                        }
                    )

                    if busqueda:
                        registros_omitidos += 1
                        continue

                    dict_datos = row.to_dict()
                    # Convertir posibles NaNs en None para el JSON
                    dict_datos_clean = {k: v for k, v in dict_datos.items() if pd.notna(v)}

                    create_data = {
                        "fecha": fecha,
                        "hora": hora,
                        "turno": turno,
                        "proceso": proceso,
                        "maquina": maquina,
                        "bodega": bodega,
                        "o_trabajo": o_trabajo,
                        "referencia": referencia,
                        "item_silp": item_silp,
                        "descripcion": descripcion,
                        "kilos": kilos,
                        "usuario_registra": usuario,
                        "observacion": observacion,
                        "kilos_rech": kilos_rech,
                        "rechazado_por": rechazado_por,
                        "observ_rechz": observ_rechz,
                        "estado": estado,
                        "causal": causal,
                        "datos": json.dumps(dict_datos_clean, default=str)
                    }

                    db.desperdicio.create(data=create_data)
                    registros_creados += 1

                print(f"Resumen para {archivo}: {registros_creados} nuevos registros gurdados, {registros_omitidos} registros omitidos (ya existían).")
                
                # Mover archivo a _procesados
                shutil.move(ruta_archivo, ruta_destino)
                print(f"Archivo movido a: {procesados_dir}")
                num_archivos_procesados += 1

            except Exception as e:
                print(f"Error procesando el archivo {archivo}: {e}")

        print(f"\nFinalizado: {num_archivos_procesados} archivo(s) procesado(s) existosamente.")

    finally:
        db.disconnect()

if __name__ == "__main__":
    process_desperdicios()
