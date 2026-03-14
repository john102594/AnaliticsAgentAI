import argparse
import os
import requests
import json
import urllib3

# Desactivar advertencias de certificado SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class DownloadNovedades:
    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self.config = self._load_config()
        self.base_dir = self.config.get("descargas_dir", "descargas")
        self.descargas_dir = os.path.join(self.base_dir, "novedades")
        self._crear_directorio()
        
    def _load_config(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(script_dir, self.config_path)
        with open(full_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _crear_directorio(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_dir = os.path.join(script_dir, self.descargas_dir)
        old_dir = os.path.join(full_dir, "_procesados")
        
        if not os.path.exists(full_dir):
            os.makedirs(full_dir)
        if not os.path.exists(old_dir):
            os.makedirs(old_dir)

    def descargar(self, fecha: str, turno: int):
        cred_config = self.config.get("credentials", {})
        usuario = cred_config.get("usuario")
        contraseña = cred_config.get("contraseña")
        
        # Asegurar formato fecha YYYYMMDD para la URL si viene con guiones
        fecha_url = fecha.replace("-", "")
        turno_padded = f"{turno:02d}"
        
        # Buscar config de novedades
        reportes_config = self.config.get("server", {}).get("reports", [])
        cfg = next((r for r in reportes_config if r.get("id") == "novedades"), None)
        
        if not cfg:
            print("Configuración de 'novedades' no encontrada en config.json.")
            return None

        url_pattern = cfg.get("url_pattern")
        filename_pattern = cfg.get("filename_pattern")
        
        url = url_pattern.format(fecha=fecha_url, turno=turno, turno_padded=turno_padded)
        nombre_archivo = filename_pattern.format(fecha=fecha_url, turno=turno, turno_padded=turno_padded)
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_descargas_dir = os.path.join(script_dir, self.descargas_dir)
        
        ruta_archivo = os.path.join(full_descargas_dir, nombre_archivo)
        ruta_archivo_old = os.path.join(full_descargas_dir, "_procesados", nombre_archivo)

        print(f"--- Iniciando descarga de Novedades Impresion ---")
        
        if os.path.exists(ruta_archivo):
            print(f"El archivo {nombre_archivo} ya existe en novedades: {ruta_archivo}")
            return ruta_archivo
            
        if os.path.exists(ruta_archivo_old):
            print(f"El archivo {nombre_archivo} ya existe en _procesados: {ruta_archivo_old}")
            return ruta_archivo_old

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
                return ruta_archivo
            else:
                print(f"Error al descargar Novedades Impresion. Código: {respuesta.status_code}")
                return None

        except Exception as e:
            print(f"Error procesando la solicitud para Novedades Impresion: {str(e)}")
            return None
        finally:
            if sesion:
                sesion.close()

def main():
    parser = argparse.ArgumentParser(description="Descarga reporte de Novedades Intranet Impresion.")
    parser.add_argument("--fecha", required=True, help="Fecha para el reporte (formato YYYY-MM-DD o YYYYMMDD)")
    parser.add_argument("--turno", type=int, choices=[1, 2], required=True, help="Turno (1 o 2)")
    
    args = parser.parse_args()
    
    downloader = DownloadNovedades()
    downloader.descargar(args.fecha, args.turno)

if __name__ == "__main__":
    main()
