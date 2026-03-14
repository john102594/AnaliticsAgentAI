import os
import shutil
import re
import argparse
from process_reports import ProcessReports
from process_desperdicios import process_desperdicios
from process_novedades_impresion import process_and_save_files as process_and_save_novedades
from process_tintas import process_and_save_files as process_and_save_tintas
from process_tpr import process_and_save_files as process_and_save_tpr
from process_novedades_impresion import load_config

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def process_novedades_produccion():
    print("\n=== PROCESANDO NOVEDADES DE PRODUCCIÓN ===")
    # Instanciar el procesador
    processor = ProcessReports()
    
    # Obtener configuración de directorios
    descargas_dir = processor.config.get("descargas_dir", "descargas")
    full_descargas_dir = os.path.join(BASE_DIR, descargas_dir, "reporteturno")
    
    # Directorio de históricos procesados
    old_dir = os.path.join(full_descargas_dir, "_procesados")
    if not os.path.exists(old_dir):
        os.makedirs(old_dir)
        print(f"Creado directorio: {old_dir}")

    # Patrón para identificar archivos: Novedades_Turno_{turno}_{fecha}.xlsx
    pattern = re.compile(r"Novedades_Turno_(\d)_(\d{8})\.xlsx")

    # Listar archivos en descargas
    if not os.path.exists(full_descargas_dir):
        print(f"El directorio {full_descargas_dir} no existe.")
        return

    archivos = [f for f in os.listdir(full_descargas_dir) if os.path.isfile(os.path.join(full_descargas_dir, f))]
    
    archivos_validos = [f for f in archivos if pattern.match(f)]
    
    if not archivos_validos:
        print("No se encontraron archivos de producción para procesar en descargas.")
        return

    print(f"Encontrados {len(archivos_validos)} archivos de producción. Iniciando procesamiento...")

    fechas_procesadas = []

    for archivo in archivos_validos:
        match = pattern.match(archivo)
        turno = int(match.group(1))
        fecha_bruta = match.group(2) # YYYYMMDD
        fecha_iso = f"{fecha_bruta[:4]}-{fecha_bruta[4:6]}-{fecha_bruta[6:]}"
        fechas_procesadas.append(fecha_iso)
        
        ruta_origen = os.path.join(full_descargas_dir, archivo)
        ruta_destino = os.path.join(old_dir, archivo)

        print(f"\n--- Procesando: {archivo} (Fecha: {fecha_iso}, Turno: {turno}) ---")
        
        try:
            processor.procesar(fecha_iso, turno)
            shutil.move(ruta_origen, ruta_destino)
            print(f"Archivo movido exitosamente a: {old_dir}")
            
        except Exception as e:
            print(f"Error crítico procesando {archivo}: {e}")

    # Post-procesamiento de deltas para mejorar la velocidad
    if fechas_procesadas:
        fecha_minima = min(fechas_procesadas)
        print(f"\nIniciando cálculo masivo de deltas desde {fecha_minima}...")
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
