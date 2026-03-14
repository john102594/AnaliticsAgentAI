import argparse
import os
from datetime import datetime, timedelta
from download_reports import DownloadReports
from download_desperdicios import DownloadDesperdicios
from download_novedades import DownloadNovedades
from download_tintas import DownloadTintas
from download_tpr import DownloadTPR

def batch_download(tipo_reporte, fecha_inicio_str=None, fecha_fin_str=None):
    # Rango de fechas
    if fecha_inicio_str:
        fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
    else:
        fecha_inicio = datetime(2026, 3, 1)
        
    if fecha_fin_str:
        fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d")
    else:
        fecha_fin = datetime.now()
    
    # Turnos a descargar
    turnos = [1, 2]
    
    # Instanciar descargadores según selección
    downloader_reportes = DownloadReports() if tipo_reporte in ["novedades", "reportes", "todos"] else None
    downloader_desperdicios = DownloadDesperdicios() if tipo_reporte in ["desperdicios", "todos"] else None
    downloader_novedades_imp = DownloadNovedades() if tipo_reporte in ["novedades_impresion", "todos"] else None
    downloader_tintas = DownloadTintas() if tipo_reporte in ["tintas","diario", "semanal"] else None
    downloader_tpr = DownloadTPR() if tipo_reporte in ["tpr", "diario", "semanal"] else None
    
    
    print(f"Iniciando descarga masiva de '{tipo_reporte}' desde {fecha_inicio.strftime('%Y-%m-%d')} hasta {fecha_fin.strftime('%Y-%m-%d')}")
    
    current_date = fecha_inicio
    while current_date <= fecha_fin:
        fecha_str = current_date.strftime("%Y%m%d")
        print(f"\n--- Procesando fecha: {fecha_str} ---")
        
        for turno in turnos:
            print(f"Turno: {turno}")
            
            if downloader_reportes:
                try:
                    downloader_reportes.descargar_reportes(fecha_str, turno)
                except Exception as e:
                    print(f"Error descargando reportes para fecha {fecha_str} turno {turno}: {e}")
            
            if downloader_desperdicios:
                try:
                    downloader_desperdicios.descargar(fecha_str, turno)
                except Exception as e:
                    print(f"Error descargando desperdicios para fecha {fecha_str} turno {turno}: {e}")
                    
            if downloader_novedades_imp:
                try:
                    downloader_novedades_imp.descargar(fecha_str, turno)
                except Exception as e:
                    print(f"Error descargando novedades impresion para fecha {fecha_str} turno {turno}: {e}")
                    
        # Tintas is accumulated daily, not by shift.
        if downloader_tintas:
            try:
                downloader_tintas.descargar(fecha_str)
            except Exception as e:
                print(f"Error descargando tintas para fecha {fecha_str}: {e}")
                
        # TPR is daily
        if downloader_tpr:
            try:
                downloader_tpr.descargar(fecha_str)
            except Exception as e:
                print(f"Error descargando tpr para fecha {fecha_str}: {e}")
                
        current_date += timedelta(days=1)
        
    print("\nProceso de descarga masiva finalizado.")

def main():
    parser = argparse.ArgumentParser(description="Realiza descarga masiva de reportes en un rango de fechas predefinido.")
    parser.add_argument(
        "--tipo", 
        choices=["novedades", "desperdicios", "novedades_impresion", "tintas", "tpr", "diario", "semanal", "todos"], 
        default="todos",
        help="Tipo de reporte a descargar: novedades, desperdicios, novedades_impresion, tintas, tpr, diario, semanal o todos (por defecto)."
    )
    parser.add_argument(
        "--fecha_inicio", 
        type=str, 
        help="Fecha inicial para descarga (YYYY-MM-DD)."
    )
    parser.add_argument(
        "--fecha_fin", 
        type=str, 
        help="Fecha final para descarga (YYYY-MM-DD)."
    )
    args = parser.parse_args()
    
    batch_download(args.tipo, args.fecha_inicio, args.fecha_fin)

if __name__ == "__main__":
    main()
