import argparse
from process_reports import ProcessReports

def run_recalc():
    parser = argparse.ArgumentParser(description="Recalcula los deltas de la tabla Historico.")
    parser.add_argument("--desde", default="2026-02-01", help="Fecha desde la cual recalcular (YYYY-MM-DD)")
    
    args = parser.parse_args()
    
    processor = ProcessReports()
    print(f"Iniciando recalculo de deltas desde: {args.desde}")
    
    try:
        processor.post_procesar_deltas(args.desde)
        print("Recalculo completado exitosamente.")
    except Exception as e:
        print(f"Error durante el recalculo: {e}")

if __name__ == "__main__":
    run_recalc()
