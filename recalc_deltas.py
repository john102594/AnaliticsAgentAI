import argparse
import sys
from prisma.errors import PrismaError
from process_reports import ProcessReports

def run_recalc():
    parser = argparse.ArgumentParser(description="Recalcula los deltas de la tabla Historico.")
    parser.add_argument("--desde", default="2026-02-01", help="Fecha desde la cual recalcular (YYYY-MM-DD)")
    parser.add_argument("--hasta", default=None, help="Fecha hasta la cual recalcular (YYYY-MM-DD)")
    
    args = parser.parse_args()
    
    processor = ProcessReports()
    
    try:
        processor.post_procesar_deltas(args.desde, args.hasta)
        print("Recalculo completado exitosamente.")
    except (ValueError, PrismaError) as e:
        print(f"Error durante el recalculo: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run_recalc()
