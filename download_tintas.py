import argparse
import os
import requests
import json
import urllib3

# Desactivar advertencias de certificado SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class DownloadTintas:
    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self.config = self._load_config()
        self.descargas_dir = os.path.join(self.config.get("descargas_dir", "descargas"), "tintas")
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

    def descargar(self, fecha: str):
        cred_config = self.config.get("credentials", {})
        usuario = cred_config.get("usuario")
        contraseña = cred_config.get("contraseña")
        
        # Formato YYYYMMDD  
        fecha_url = fecha.replace("-", "")

        # Buscar config de tintas
        reportes_config = self.config.get("server", {}).get("reports", [])
        cfg = next((r for r in reportes_config if r.get("id") == "tintas"), None)
        
        if not cfg:
            print("Configuración de 'tintas' no encontrada en config.json.")
            return []

        url = cfg.get("url_pattern").format(fecha=fecha_url)
        nombre_archivo = cfg.get("filename_pattern").format(fecha=fecha_url)
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_descargas_dir = os.path.join(script_dir, self.descargas_dir)

        ruta_archivo = os.path.join(full_descargas_dir, nombre_archivo)
        ruta_archivo_proc = os.path.join(full_descargas_dir, "_procesados", nombre_archivo)

        print(f"--- Iniciando descarga de Consumo de Tintas ---")
        
        if os.path.exists(ruta_archivo):
            print(f"El archivo {nombre_archivo} ya existe: {ruta_archivo}")
            return [ruta_archivo]
            
        if os.path.exists(ruta_archivo_proc):
            print(f"El archivo {nombre_archivo} ya existe en _procesados: {ruta_archivo_proc}")
            return [ruta_archivo_proc]

        print(f"Descargando desde: {url}")
        
        sesion = None
        descargados = []
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
                descargados.append(ruta_archivo)
            else:
                print(f"Error al descargar. Código: {respuesta.status_code}")

        except Exception as e:
            print(f"Error procesando la solicitud para Tintas: {str(e)}")
        finally:
            if sesion:
                sesion.close()

        return descargados

def main():
    parser = argparse.ArgumentParser(description="Descarga reportes de Consumo de Tintas.")
    parser.add_argument("--fecha", required=True, help="Fecha para el reporte (formato YYYY-MM-DD o YYYYMMDD)")
    
    args = parser.parse_args()
    
    downloader = DownloadTintas()
    downloader.descargar(args.fecha)

if __name__ == "__main__":
    main()
