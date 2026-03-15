import os
import shutil
import re
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from process_reports import ProcessReports
from process_desperdicios import process_desperdicios
from process_novedades_impresion import process_and_save_files as process_and_save_novedades
from process_tintas import process_and_save_files as process_and_save_tintas
from process_tpr import process_and_save_files as process_and_save_tpr
from process_novedades_impresion import load_config

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def process_novedades_produccion():
    print("\n=== PROCESANDO NOVEDADES DE PRODUCCIÓN ===")
    processor = ProcessReports()
    
    descargas_dir = processor.config.get("descargas_dir", "descargas")
    full_descargas_dir = os.path.join(BASE_DIR, descargas_dir, "reporteturno")
    
    old_dir = os.path.join(full_descargas_dir, "_procesados")
    if not os.path.exists(old_dir):
        os.makedirs(old_dir)
        print(f"Creado directorio: {old_dir}")

    pattern = re.compile(r"Novedades_Turno_(\d)_(\d{8})\.xlsx")

    if not os.path.exists(full_descargas_dir):
        print(f"El directorio {full_descargas_dir} no existe.")
        return

    archivos = [f for f in os.listdir(full_descargas_dir) if os.path.isfile(os.path.join(full_descargas_dir, f))]
    archivos_validos = sorted([f for f in archivos if pattern.match(f)])
    
    if not archivos_validos:
        print("No se encontraron archivos de producción para procesar en descargas.")
        return

    print(f"Encontrados {len(archivos_validos)} archivos. Fase 1: verificando existencia en BD...")

    # --- FASE 1: Verificar cuáles ya existen en BD y filtrar solo los nuevos ---
    processor.db.connect()
    try:
        archivos_a_procesar = []
        for archivo in archivos_validos:
            match = pattern.match(archivo)
            turno = int(match.group(1))
            fecha_bruta = match.group(2)
            fecha_iso = f"{fecha_bruta[:4]}-{fecha_bruta[4:6]}-{fecha_bruta[6:]}"
            
            # Calcular fecha/turno real en BD (igual que en procesar())
            turno_db = turno
            fecha_db = fecha_iso
            if turno == 2:
                import pandas as pd_local
                dt = pd_local.to_datetime(fecha_iso)
                if dt.day == 1:
                    turno_db = 0
                else:
                    fecha_db = (dt - pd_local.Timedelta(days=1)).strftime('%Y-%m-%d')
            
            ya_existe = processor.db.historico.find_first(where={"fecha": fecha_db, "turno": turno_db})
            if ya_existe:
                print(f"  [SKIP] {archivo} → ya en BD ({fecha_db} T{turno_db})")
                # Mover a procesados aunque ya esté en BD
                ruta_origen = os.path.join(full_descargas_dir, archivo)
                ruta_destino = os.path.join(old_dir, archivo)
                if os.path.exists(ruta_origen):
                    shutil.move(ruta_origen, ruta_destino)
            else:
                archivos_a_procesar.append((archivo, fecha_iso, turno))
    finally:
        processor.db.disconnect()

    if not archivos_a_procesar:
        print("Todos los archivos ya están en la BD. Nada que procesar.")
        return

    print(f"\n{len(archivos_a_procesar)} archivos nuevos. Fase 2: parsing paralelo de Excel...")

    # --- FASE 2: Parsear todos los Excel EN PARALELO (sin tocar la BD) ---
    def _parsear_archivo(args):
        """Solo lee el Excel y devuelve los datos parseados. Sin escritura a BD."""
        archivo, fecha_iso, turno = args
        ruta = os.path.join(full_descargas_dir, archivo)
        try:
            datos = processor.parsear_excel(ruta, fecha_iso, turno)
            return archivo, fecha_iso, turno, datos, None
        except Exception as e:
            return archivo, fecha_iso, turno, None, str(e)

    MAX_WORKERS = 4
    resultados_parseados = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_parsear_archivo, args): args[0] for args in archivos_a_procesar}
        for future in as_completed(futures):
            archivo, fecha_iso, turno, datos, error = future.result()
            if error:
                print(f"  [ERROR parse] {archivo}: {error}")
            else:
                resultados_parseados.append((archivo, fecha_iso, turno, datos))

    # Ordenar por fecha+turno para consistencia cronológica
    resultados_parseados.sort(key=lambda x: (x[1], x[2]))

    # --- FASE 3: Acumular TODOS los registros y hacer un solo bulk insert ---
    print(f"\nFase 3: acumulando registros de {len(resultados_parseados)} archivos para bulk insert...")

    all_historico = []
    all_novedad = []

    for archivo, fecha_iso, turno, datos in resultados_parseados:
        all_novedad.extend(datos.get("novedades", []))
        for filas in datos.get("historico", {}).values():
            all_historico.extend(filas)

    print(f"  Total: {len(all_historico)} registros historico, {len(all_novedad)} novedades.")
    print("  Insertando en BD...")

    fechas_procesadas = []
    processor.db.connect()
    try:
        if all_historico:
            processor.db.historico.create_many(data=all_historico)
        if all_novedad:
            processor.db.novedad.create_many(data=all_novedad)

        # Mover archivos procesados exitosamente
        for archivo, fecha_iso, turno, datos in resultados_parseados:
            ruta_origen = os.path.join(full_descargas_dir, archivo)
            ruta_destino = os.path.join(old_dir, archivo)
            if os.path.exists(ruta_origen):
                shutil.move(ruta_origen, ruta_destino)
            fechas_procesadas.append(fecha_iso)

        print(f"  ✓ Bulk insert completado.")
    except Exception as e:
        print(f"  [ERROR bulk insert] {e}")
    finally:
        processor.db.disconnect()


    # Post-procesamiento de deltas
    if fechas_procesadas:
        fecha_minima = min(fechas_procesadas)
        print(f"\nIniciando cálculo de deltas desde {fecha_minima}...")
        try:
            processor.post_procesar_deltas(fecha_minima)
        except Exception as e:
            print(f"Error calculando deltas: {e}")






def process_desperdicios_wrapper():
    print("\n=== PROCESANDO DESPERDICIOS ===")
    process_desperdicios()


def process_novedades_impresion_wrapper():
    print("\n=== PROCESANDO NOVEDADES DE IMPRESIÓN ===")
    config = load_config()
    if not config:
        print("No se pudo cargar config.json")
        return
        
    target_dir = os.path.join(BASE_DIR, config.get("download_dir", "descargas"), "novedades")
    files_to_process = []
    
    if os.path.exists(target_dir):
        for filename in os.listdir(target_dir):
            if filename.endswith(".xls") or filename.endswith(".xlsx"):
                file_path = os.path.join(target_dir, filename)
                if os.path.isfile(file_path):
                    files_to_process.append(file_path)
        
        if files_to_process:
            process_and_save_novedades(files_to_process)
        else:
            print(f"No hay archivos de Novedades de Impresión para procesar en: {target_dir}")
    else:
        print(f"El directorio no existe: {target_dir}")


def process_tintas_wrapper():
    print("\n=== PROCESANDO TINTAS ===")
    config = load_config()
    if not config:
        print("No se pudo cargar config.json")
        return
        
    target_dir = os.path.join(BASE_DIR, config.get("download_dir", "descargas"), "tintas")
    
    if os.path.exists(target_dir):
        # We process all of them
        process_and_save_tintas()
    else:
        print(f"El directorio no existe o no tiene archivos de tintas: {target_dir}")


def process_tpr_wrapper():
    print("\n=== PROCESANDO TPR ===")
    config = load_config()
    if not config:
        print("No se pudo cargar config.json")
        return
        
    target_dir = os.path.join(BASE_DIR, config.get("download_dir", "descargas"), "TPR")
    
    if os.path.exists(target_dir):
        process_and_save_tpr()
    else:
        print(f"El directorio no existe o no tiene archivos de TPR: {target_dir}")


def batch_process(tipo_reporte):
    print(f"Iniciando procesamiento masivo de tipo: '{tipo_reporte}'")
    
    if tipo_reporte in ["novedades_turno", "produccion", "todos"]:
        process_novedades_produccion()
        
    if tipo_reporte in ["desperdicios", "todos"]:
        process_desperdicios_wrapper()
        
    if tipo_reporte in ["novedades_impresion", "todos"]:
        process_novedades_impresion_wrapper()
        
    if tipo_reporte in ["tintas", "diario", "semanal"]:
        process_tintas_wrapper()
        
    if tipo_reporte in ["tpr", "diario", "semanal"]:
        process_tpr_wrapper()
        
    print("\nProcesamiento masivo finalizado.")


def main():
    parser = argparse.ArgumentParser(description="Procesa multiples reportes descargados desde sus carpetas correspondientes.")
    parser.add_argument(
        "--tipo", 
        choices=["novedades_turno", "desperdicios", "novedades_impresion", "tintas", "tpr", "diario", "semanal", "todos"], 
        default="todos",
        help="Tipo de reporte a procesar (por defecto: todos)."
    )
    args = parser.parse_args()
    
    batch_process(args.tipo)

if __name__ == "__main__":
    main()
