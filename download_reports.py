import argparse
import os
import requests
import json
import urllib3
from datetime import datetime

# Desactivar advertencias de certificado SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class DownloadReports:
    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self.config = self._load_config()
        self.descargas_dir = os.path.join(self.config.get("descargas_dir", "descargas"), "reporteturno")
        self._crear_directorio()
        
    def _load_config(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(script_dir, self.config_path)
        with open(full_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _crear_directorio(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_dir = os.path.join(script_dir, self.descargas_dir)
        if not os.path.exists(full_dir):
            os.makedirs(full_dir)

    def descargar_reportes(self, fecha: str, turno: int):
        cred_config = self.config.get("credentials", {})
        usuario = cred_config.get("usuario")
        contraseña = cred_config.get("contraseña")
        
        reportes_config = self.config.get("server", {}).get("reports", [])
        
        if not reportes_config:
            print("No hay reportes configurados en config.json.")
            return

        # Asegurar formato fecha YYYYMMDD para la URL si viene con guiones
        fecha_url = fecha.replace("-", "")
        turno_str = str(turno)
        turno_padded = f"{turno:02d}"

        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_descargas_dir = os.path.join(script_dir, self.descargas_dir)

        Archivos_descargados = []

        for reporte in reportes_config:
            if reporte.get("id") != "reporteturno":
                continue
                
            nombre_reporte = reporte.get('name', 'Reporte')
            url_pattern = reporte.get('url_pattern')
            filename_pattern = reporte.get('filename_pattern')

            if not url_pattern or not filename_pattern:
                print(f"Configuración incompleta para reporte: {nombre_reporte}")
                continue

            url = url_pattern.format(fecha=fecha_url, turno=turno, turno_padded=turno_padded)
            # Para el nombre de archivo mantenemos el formato original
            nombre_archivo = filename_pattern.format(fecha=fecha_url, turno=turno, turno_padded=turno_padded)
            ruta_archivo = os.path.join(full_descargas_dir, nombre_archivo)
            ruta_archivo_old = os.path.join(full_descargas_dir, "_procesados", nombre_archivo)

            print(f"--- Iniciando descarga de {nombre_reporte} ---")
            
            if os.path.exists(ruta_archivo):
                print(f"El archivo {nombre_archivo} ya existe: {ruta_archivo}")
                Archivos_descargados.append(ruta_archivo)
                continue
                
            if os.path.exists(ruta_archivo_old):
                print(f"El archivo {nombre_archivo} ya existe en _procesados: {ruta_archivo_old}")
                Archivos_descargados.append(ruta_archivo_old)
                continue

            print(f"Descargando desde: {url}")
            
            sesion = None
            try:
                sesion = requests.Session()
                sesion.verify = False

                respuesta = sesion.get(
                    url,
                    auth=(usuario, contraseña),
                    timeout=30
                )

                if respuesta.status_code == 200:
                    with open(ruta_archivo, 'wb') as f:
                        f.write(respuesta.content)
                    print(f"Archivo descargado exitosamente: {ruta_archivo}")
                    Archivos_descargados.append(ruta_archivo)
                else:
                    print(f"Error al descargar {nombre_reporte}. Código: {respuesta.status_code}")

            except Exception as e:
                print(f"Error procesando la solicitud para {nombre_reporte}: {str(e)}")
            finally:
                if sesion:
                    sesion.close()

        return Archivos_descargados


def main():
    parser = argparse.ArgumentParser(description="Descarga reportes de producción.")
    parser.add_argument("--fecha", required=True, help="Fecha para el reporte (formato YYYY-MM-DD o YYYYMMDD)")
    parser.add_argument("--turno", type=int, choices=[1, 2], required=True, help="Turno (1 o 2)")
    
    args = parser.parse_args()
    
    downloader = DownloadReports()
    downloader.descargar_reportes(args.fecha, args.turno)

if __name__ == "__main__":
    main()
