"""
Module for processing Excel production reports and synchronizing data with the database.

This script extracts data from 'Novedades' and 'Históricos' Excel files, cleans and maps 
the information, and saves it into the database using Prisma. It also includes 
functionality to calculate production deltas between consecutive shifts.
"""
import os
import json
import argparse
import pandas as pd
from prisma import Prisma
from prisma.errors import PrismaError

class ProcessReports:
    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self._load_config()
        self.db = Prisma()
        
    def _load_config(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(script_dir, self.config_path)
        with open(full_path, "r", encoding="utf-8") as f:
            self.config = json.load(f)
            
    # ---- LOGICA EXTRAIDA DE excel_processor.py (NOVEDADES) ----
    def _find_table_start(self, df):
        for idx in range(len(df)):
            valor = df.iloc[idx, 0]
            if isinstance(valor, str) and "Detallado de Novedades" in valor:
                return idx + 2, df.iloc[idx + 2, 1]
        raise ValueError("No se encontro el inicio de la tabla de novedades")
        
    def _find_table_end(self, df, inicio):
        for idx in range(inicio, len(df)):
            valor = df.iloc[idx, 0]
            if isinstance(valor, str) and "Detalle de Rechazos" in valor:
                return idx
        return len(df)

    def _process_table_novedades(self, df, inicio, fin, primera_maquina, fecha, turno):
        tabla = df.iloc[inicio:fin].reset_index(drop=True)
        if tabla.empty or len(tabla) < 2:
            return pd.DataFrame()

        encabezados = tabla.iloc[1]
        tabla = tabla.iloc[2:].reset_index(drop=True)
        tabla.columns = [str(col).strip() for col in encabezados]
        
        filas_vacias = tabla[tabla.iloc[:, 0].isna()].index.tolist()
        filas_vacias.append(len(tabla))
        
        tablas_procesadas = []
        inicio_idx = 0
        es_primera_seccion = True
        
        if filas_vacias and filas_vacias[0] == 0:
            es_primera_seccion = False
            
        for fin_seccion in filas_vacias:
            if fin_seccion >= inicio_idx:
                seccion = tabla.iloc[inicio_idx:fin_seccion].copy()
                if not seccion.empty:
                    if es_primera_seccion:
                        nombre_maquina = primera_maquina
                        es_primera_seccion = False
                    else:
                        nombre_maquina = seccion.iloc[0, 1]
                        seccion = seccion.iloc[1:]
                    
                    if pd.notna(nombre_maquina) and str(nombre_maquina).strip() != '':
                        nombre_maquina = str(nombre_maquina).strip()
                        if not seccion.empty:
                            seccion.insert(0, 'Turno', turno)
                            seccion.insert(0, 'Fecha', fecha)
                            seccion.insert(0, 'Maquina', nombre_maquina)
                            tablas_procesadas.append(seccion)
                inicio_idx = fin_seccion + 1
                
        if not tablas_procesadas:
            return pd.DataFrame()
            
        tabla_final = pd.concat(tablas_procesadas, ignore_index=True)
        tabla_final = tabla_final.dropna(axis=1, how='all')
        tabla_final = tabla_final[tabla_final['Refer'] != 'Refer']
        
        columnas_numericas = ['Order', 'ESPEAUT', 'ESPEPRO', 'ESPEOPE', 'Cantidad']
        for col in columnas_numericas:
            if col in tabla_final.columns:
                tabla_final[col] = pd.to_numeric(tabla_final[col], errors='coerce')
                
        return tabla_final

    # ---- LOGICA EXTRAIDA DE leer_reporte.py (HISTORICOS) ----
    def _extraer_tabla_historico(self, df):
        patron = r'Cumplimiento de Producci[oó]n y Desperdicio Acumulado'
        mask = df.iloc[:, 0].astype(str).str.contains(patron, na=False)
        if not mask.any():
            return None
        inicio = df[mask].index[0]
        
        mask_fin = df.iloc[inicio:].iloc[:, 0].astype(str).str.contains('TOTAL', na=False)
        if not mask_fin.any():
            return None
        fin = df.iloc[inicio:][mask_fin].index[0]
        
        tabla = df.iloc[inicio+1:fin+1].copy()
        encabezados = tabla.iloc[0]
        tabla = tabla.iloc[1:]
        tabla.columns = encabezados
        
        return tabla

    def _limpiar_tabla_historico(self, tabla, fecha, turno):
        tabla.rename(columns={tabla.columns[0]: 'Maquina'}, inplace=True)
        tabla = tabla.loc[:, ~tabla.columns.isna()]
        tabla.insert(0, 'Fecha', fecha)
        tabla.insert(1, 'Turno', turno)
        tabla = tabla.iloc[:-1] # Remove TOTAL row
        tabla.reset_index(drop=True, inplace=True)
        return tabla

    def _safe_float(self, val):
        if pd.isna(val) or val is None:
            return None
        try:
            # Eliminar '%' y espacios si vienen
            val_str = str(val).replace('%', '').replace(',', '.').strip()
            # Si quedo vacio
            if not val_str:
                return None
            return float(val_str)
        except ValueError:
            return None

    def _get_val(self, row_dict, *keys):
        for k in keys:
            if k in row_dict:
                val = row_dict[k]
                if pd.notna(val):
                    return val
        return None

    def _safe_str(self, val):
        if pd.isna(val) or val is None:
            return None
        return str(val).strip()

    def _map_novedad_dict(self, row_dict):
        mapped = {}
        # Strings
        mapped['orden'] = self._safe_str(self._get_val(row_dict, 'Orden', 'Order'))
        mapped['refer'] = self._safe_str(self._get_val(row_dict, 'Refer'))
        mapped['cump_horas'] = self._safe_float(self._get_val(row_dict, 'Cump.en Horas', 'Cump.en.Horas'))
        mapped['cump_metros'] = self._safe_float(self._get_val(row_dict, 'Cump.en Metros', 'Cump.en.Metros'))
        mapped['cump_t1_t4'] = self._safe_float(self._get_val(row_dict, 'Cump.T1-T4'))
        mapped['kg_prod'] = self._safe_float(self._get_val(row_dict, 'Kg.Prod.'))
        mapped['kg_prog'] = self._safe_float(self._get_val(row_dict, 'Kg.Prog.'))
        mapped['mts_prod'] = self._safe_float(self._get_val(row_dict, 'Mts.Prod.'))
        mapped['mts_prog'] = self._safe_float(self._get_val(row_dict, 'Mts.Prog.'))
        mapped['desp_no_r_real'] = self._safe_float(self._get_val(row_dict, 'Desp.No R Real'))
        mapped['desp_no_r_std'] = self._safe_float(self._get_val(row_dict, 'Desp.No R Std.'))
        mapped['mts_no_r_real'] = self._safe_float(self._get_val(row_dict, 'Mts.No R Real'))
        mapped['mts_no_r_std'] = self._safe_float(self._get_val(row_dict, 'Mts.No R Std.'))
        mapped['desp_r_real'] = self._safe_float(self._get_val(row_dict, 'Desp. R Real'))
        mapped['t1_real'] = self._safe_float(self._get_val(row_dict, 'T1 Real', 'T1.Real'))
        mapped['t1_std'] = self._safe_float(self._get_val(row_dict, 'T1 Std.', 'T1.Std.'))
        mapped['t2_real'] = self._safe_float(self._get_val(row_dict, 'T2 Real', 'T2.Real'))
        mapped['t2_std'] = self._safe_float(self._get_val(row_dict, 'T2 Std.', 'T2.Std.'))
        mapped['t3_real'] = self._safe_float(self._get_val(row_dict, 'T3 Real', 'T3.Real'))
        mapped['t3_std'] = self._safe_float(self._get_val(row_dict, 'T3 Std.', 'T3.Std.'))
        mapped['t4_real'] = self._safe_float(self._get_val(row_dict, 'T4 Real', 'T4.Real'))
        mapped['t4_std'] = self._safe_float(self._get_val(row_dict, 'T4 Std.', 'T4.Std.'))
        mapped['t5_real'] = self._safe_float(self._get_val(row_dict, 'T5 Real', 'T5.Real'))
        mapped['t5_std'] = self._safe_float(self._get_val(row_dict, 'T5 Std.', 'T5.Std.'))
        mapped['vel_real'] = self._safe_float(self._get_val(row_dict, 'Vel.Real'))
        mapped['vel_std'] = self._safe_float(self._get_val(row_dict, 'Vel.Std.'))
        mapped['t_parada'] = self._safe_float(self._get_val(row_dict, 'T Parada', 'T.Parada'))
        mapped['num_parada'] = self._safe_float(self._get_val(row_dict, 'Num.Parada'))
        mapped['mts_por_bob_prod'] = self._safe_float(self._get_val(row_dict, 'Mts x Bob Prod.', 'Mts.x.Bob.Prod.'))
        mapped['kg_por_bob_prod'] = self._safe_float(self._get_val(row_dict, 'Kg x Bob Prod', 'Kg.x.Bob.Prod.', 'Kg_x_Bob_Prod'))
        mapped['nro_bob_prod'] = self._safe_float(self._get_val(row_dict, 'Nro Bob Prod', 'Nro.Bob.Prod.'))
        mapped['t_total_real'] = self._safe_float(self._get_val(row_dict, 'T Total Real', 'T.Total.Real'))
        mapped['t_total_std'] = self._safe_float(self._get_val(row_dict, 'T Total Std.', 'T.Total.Std.'))
        mapped['productividad'] = self._safe_float(self._get_val(row_dict, 'Productividad'))
        mapped['paradas_std_t2'] = self._safe_float(self._get_val(row_dict, '# Paradas Estandar T2'))
        mapped['paradas_reales_t2'] = self._safe_float(self._get_val(row_dict, '# Paradas Reales T2'))
        mapped['paradas_std_t3'] = self._safe_float(self._get_val(row_dict, '# Paradas Estandar T3'))
        mapped['paradas_reales_t3'] = self._safe_float(self._get_val(row_dict, '# Paradas Reales T3'))
        mapped['paradas_std_t4'] = self._safe_float(self._get_val(row_dict, '# Paradas Estandar T4'))
        mapped['paradas_reales_t4'] = self._safe_float(self._get_val(row_dict, '# Paradas Reales T4'))
        mapped['paradas_std_t5'] = self._safe_float(self._get_val(row_dict, '# Paradas Estandar T5'))
        mapped['paradas_reales_t5'] = self._safe_float(self._get_val(row_dict, '# Paradas Reales T5'))
        mapped['paradas_std'] = self._safe_float(self._get_val(row_dict, '# Paradas Estandar'))
        mapped['paradas_reales'] = self._safe_float(self._get_val(row_dict, '# Paradas Reales'))

        return {k: v for k, v in mapped.items() if v is not None}

    def _map_historico_dict(self, row_dict):
        # Mapea los nombres inconsistentes de las columnas de excel a nuestros campos de BD explícitos
        # Prioridad para Mapeos Alternativos (ej Impresion vs Laminacion)
        
        mapped = {}
        # Cump Horas
        mapped['cump_horas'] = self._safe_float(row_dict.get('Cump.en Horas') or row_dict.get('Cump.en.Horas'))
        mapped['cump_metros'] = self._safe_float(row_dict.get('Cump.en Metros') or row_dict.get('Cump.en.Metros'))
        mapped['cump_desperdicio_pct'] = self._safe_float(row_dict.get('Cump.Desperdicio(%)'))
        mapped['cump_productividad_pct'] = self._safe_float(row_dict.get('Cump.productividad(%)') or row_dict.get('Cump.Productividad(%)'))
        
        mapped['prod_kg'] = self._safe_float(row_dict.get('Prod.(Kg.)'))
        mapped['prod_metros'] = self._safe_float(row_dict.get('Prod.(Metros)') or row_dict.get('Prod.(Mts.)'))
        mapped['mts_std'] = self._safe_float(row_dict.get('Mts.Std.'))
        mapped['mts_cargue'] = self._safe_float(row_dict.get('Mts.Cargue'))
        mapped['kls_desperdicio'] = self._safe_float(row_dict.get('Kls.Desperdicio'))
        
        # Desperdicios variables
        mapped['desp_real_pct'] = self._safe_float(row_dict.get('% Desp. Real'))
        mapped['desp_std_pct'] = self._safe_float(row_dict.get('% Desp. Std'))
        mapped['rechazo_real_pct'] = self._safe_float(row_dict.get('% Rechazo Real'))
        mapped['kls_rechazo'] = self._safe_float(row_dict.get('Kls.Rechazo'))
        mapped['kls_refile'] = self._safe_float(row_dict.get('Kls.Refile'))
        mapped['desperdicio'] = self._safe_float(row_dict.get('Desp.') or row_dict.get('Desperdicio'))
        mapped['std_desp'] = self._safe_float(row_dict.get('Std.Desp.'))
        mapped['rechazos_kg'] = self._safe_float(row_dict.get('Rechazos(Kg.)'))
        mapped['rechazos_mts'] = self._safe_float(row_dict.get('Rechazos(Mts.)'))
        mapped['desp_r_real'] = self._safe_float(row_dict.get('Desp. R. Real'))
        mapped['desp_pct'] = self._safe_float(row_dict.get('% Desp.'))
        
        # Variaciones (impresion usa Var Cump, y los otros Var.Horas)
        mapped['var_horas'] = self._safe_float(row_dict.get('Var.Horas') or row_dict.get('Var.Cump.'))
        mapped['var_t1_t4'] = self._safe_float(row_dict.get('Var. T1-T4'))
        mapped['var_t1'] = self._safe_float(row_dict.get('Var.T1'))
        mapped['var_t2'] = self._safe_float(row_dict.get('Var.T2'))
        mapped['var_t3'] = self._safe_float(row_dict.get('Var.T3'))
        mapped['var_t4'] = self._safe_float(row_dict.get('Var.T4'))
        mapped['var_t5'] = self._safe_float(row_dict.get('Var.T5'))
        mapped['t9'] = self._safe_float(row_dict.get('T9'))
        mapped['t_sin_liq'] = self._safe_float(row_dict.get('T.sin Liq.'))
        
        # Limpiar nulos del dict para no mandar dicts con keys en None a prisma
        return {k: v for k, v in mapped.items() if v is not None}

    # ---- PROCESAMIENTO PRINCIPAL ----

    def _calcular_fecha_turno_db(self, fecha: str, turno: int):
        """Calcula la fecha y turno tal como se almacenan en la BD."""
        turno_db = turno
        fecha_db = fecha
        if turno == 2:
            dt_fecha = pd.to_datetime(fecha)
            if dt_fecha.day == 1:
                turno_db = 0
            else:
                fecha_db = (dt_fecha - pd.Timedelta(days=1)).strftime('%Y-%m-%d')
        return fecha_db, turno_db

    def parsear_excel(self, ruta_archivo: str, fecha: str, turno: int) -> dict:
        """Fase 1 (thread-safe): lee el Excel y devuelve los datos listos para guardar.
        
        No toca la BD. Puede ejecutarse en paralelo desde múltiples threads.
        Retorna un dict con 'fecha_db', 'turno_db', 'novedades' e 'historico'.
        """
        fecha_db, turno_db = self._calcular_fecha_turno_db(fecha, turno)

        reports_cfg = self.config.get("server", {}).get("reports", [])
        reporte_cfg = next((r for r in reports_cfg if r.get("id") == "reporteturno"), {})

        resultado = {"fecha_db": fecha_db, "turno_db": turno_db, "novedades": [], "historico": {}}

        # Abrir workbook UNA sola vez para todas las hojas
        xl_file = pd.ExcelFile(ruta_archivo, engine="openpyxl")

        # --- Parsear Novedades ---
        hoja_novedades = reporte_cfg.get("sheet", "IMPRESION")
        try:
            df_nov = xl_file.parse(sheet_name=hoja_novedades, header=None)
            inicio_tabla, primera_maquina = self._find_table_start(df_nov)
            fin_tabla = self._find_table_end(df_nov, inicio_tabla)
            df_novedades_final = self._process_table_novedades(df_nov, inicio_tabla, fin_tabla, primera_maquina, fecha_db, turno_db)

            for _, row in df_novedades_final.iterrows():
                row_dict = row.to_dict()
                maquina = str(row_dict.pop('Maquina', ''))
                mapped_data = self._map_novedad_dict(row_dict)
                create_data = {"fecha": fecha_db, "turno": turno_db, "maquina": maquina, "datos": json.dumps(row_dict)}
                create_data.update(mapped_data)
                resultado["novedades"].append(create_data)
        except (KeyError, IndexError) as e:
            print(f"  [WARN] Error parseando novedades en {ruta_archivo}: {e}")

        # --- Parsear Histórico ---
        hojas_historico = reporte_cfg.get("sheets_to_process", ["CORTE", "IMPRESION", "EXLAM", "LAMINACION"])
        for sheet_name in hojas_historico:
            try:
                df_hist = xl_file.parse(sheet_name=sheet_name)
                if df_hist.empty:
                    continue
                tabla_ext = self._extraer_tabla_historico(df_hist)
                if tabla_ext is None:
                    continue
                if sheet_name == "IMPRESION":
                    tabla_ext = tabla_ext.iloc[:, 1:]
                tabla_limpia = self._limpiar_tabla_historico(tabla_ext, fecha_db, turno_db)

                filas = []
                for _, row in tabla_limpia.iterrows():
                    row_dict = row.to_dict()
                    maquina = str(row_dict.pop('Maquina', ''))
                    mapped_data = self._map_historico_dict(row_dict)
                    create_data = {"fecha": fecha_db, "turno": turno_db, "maquina": maquina, "proceso": sheet_name, "datos": json.dumps(row_dict)}
                    create_data.update(mapped_data)
                    filas.append(create_data)
                resultado["historico"][sheet_name] = filas
            except ValueError:
                pass
            except (KeyError, IndexError) as e:
                print(f"  [WARN] Error parseando historico {sheet_name} en {ruta_archivo}: {e}")

        return resultado

    def guardar_datos(self, datos: dict, already_connected: bool = False):
        """Fase 2 (secuencial): escribe los datos parseados a la BD.
        
        already_connected: si True, no hace connect/disconnect.
        """
        if not already_connected:
            self.db.connect()
        try:
            fecha_db = datos["fecha_db"]
            turno_db = datos["turno_db"]

            # --- Guardar Novedades ---
            existentes_nov = self.db.novedad.find_many(where={"fecha": fecha_db, "turno": turno_db})
            used_nov_ids = set()
            count_nov_created = count_nov_updated = 0

            for create_data in datos["novedades"]:
                maquina = create_data["maquina"]
                assigned_id = next((ex.id for ex in existentes_nov if ex.maquina == maquina and ex.id not in used_nov_ids), None)
                if assigned_id:
                    used_nov_ids.add(assigned_id)
                    upd = {k: v for k, v in create_data.items() if k not in ("fecha", "turno", "maquina")}
                    self.db.novedad.update(where={"id": assigned_id}, data=upd)
                    count_nov_updated += 1
                else:
                    self.db.novedad.create(data=create_data)
                    count_nov_created += 1

            count_nov_deleted = sum(1 for ex in existentes_nov if ex.id not in used_nov_ids and not self.db.novedad.delete(where={"id": ex.id}))
            print(f"  Novedades: {count_nov_created} creados, {count_nov_updated} actualizados, {count_nov_deleted} eliminados.")

            # --- Guardar Histórico ---
            for sheet_name, filas in datos["historico"].items():
                existentes_hist = self.db.historico.find_many(where={"fecha": fecha_db, "turno": turno_db, "proceso": sheet_name})
                used_hist_ids = set()
                count_hist_created = count_hist_updated = 0

                for create_data in filas:
                    maquina = create_data["maquina"]
                    assigned_id = next((ex.id for ex in existentes_hist if ex.maquina == maquina and ex.id not in used_hist_ids), None)
                    if assigned_id:
                        used_hist_ids.add(assigned_id)
                        upd = {k: v for k, v in create_data.items() if k not in ("fecha", "turno", "maquina", "proceso")}
                        self.db.historico.update(where={"id": assigned_id}, data=upd)
                        count_hist_updated += 1
                    else:
                        self.db.historico.create(data=create_data)
                        count_hist_created += 1

                count_hist_deleted = 0
                for ex in existentes_hist:
                    if ex.id not in used_hist_ids:
                        self.db.historico.delete(where={"id": ex.id})
                        count_hist_deleted += 1
                print(f"  Histórico {sheet_name}: {count_hist_created} creados, {count_hist_updated} actualizados, {count_hist_deleted} eliminados.")
        finally:
            if not already_connected:
                self.db.disconnect()

    def procesar(self, fecha: str, turno: int, already_connected: bool = False):
        """Procesa un archivo Excel hacia la BD (compatibilidad con uso individual).
        
        Para procesamiento masivo, usar parsear_excel() + guardar_datos() por separado.
        """
        if not already_connected:
            self.db.connect()
        try:
            fecha_db, turno_db = self._calcular_fecha_turno_db(fecha, turno)
            fecha_url = fecha.replace("-", "")
            patron_archivo = f"Novedades_Turno_{turno}_{fecha_url}.xlsx"

            # Skip si ya existe
            ya_existe = self.db.historico.find_first(where={"fecha": fecha_db, "turno": turno_db})
            if ya_existe:
                print(f"  [SKIP] {patron_archivo} ya en BD ({fecha_db} T{turno_db}).")
                return True

            script_dir = os.path.dirname(os.path.abspath(__file__))
            descargas_dir = os.path.join(script_dir, self.config.get("descargas_dir", "descargas"), "reporteturno")
            ruta_archivo = os.path.join(descargas_dir, patron_archivo)

            if not os.path.exists(ruta_archivo):
                print(f"Archivo no encontrado: {ruta_archivo}")
                return

            print(f"Procesando: {ruta_archivo}")
            datos = self.parsear_excel(ruta_archivo, fecha, turno)
            self.guardar_datos(datos, already_connected=True)
        finally:
            if not already_connected:
                self.db.disconnect()




    def post_procesar_deltas(self, fecha_desde: str, fecha_hasta: str = None):
        """Calcula los deltas por turno. La suma de deltas del mes es igual al acumulado del último T1 del mes."""

        self.db.connect()
        try:
            msg = f"--- Iniciando cálculo de deltas desde {fecha_desde}"
            if fecha_hasta:
                msg += f" hasta {fecha_hasta}"
            msg += " ---"
            print(msg)

            # Construir la clausula where
            where_clause = {"fecha": {"gte": fecha_desde}}
            if fecha_hasta:
                where_clause["fecha"]["lte"] = fecha_hasta

            # Obtener registros ordenados por maquina, proceso y tiempo
            records = self.db.historico.find_many(
                where=where_clause,
                order=[
                    {"maquina": "asc"},
                    {"proceso": "asc"},
                    {"fecha": "asc"},
                    {"turno": "asc"}
                ]
            )

            print(f"Procesando {len(records)} registros...")

            def _calc_delta(curr, prev):
                if curr is None: return 0.0
                if prev is None: return float(curr)
                return max(float(curr) - float(prev), 0.0)

            # Campos cumulativos y sus campos de delta correspondientes
            FIELDS = {
                "prod_kg":      "prod_kg_turno",
                "prod_metros":  "prod_metros_turno",
                "mts_std":      "mts_std_turno",
                "mts_cargue":   "mts_cargue_turno",
                "desperdicio":  "desperdicio_turno",
                "std_desp":     "std_desp_turno",
                "rechazos_kg":  "rechazos_kg_turno",
                "rechazos_mts": "rechazos_mts_turno",
            }

            # Agrupar por (maquina, proceso) y procesar mes a mes en memoria
            from itertools import groupby
            from operator import attrgetter

            # Ordenar también por fecha/turno para el groupby
            for (maquina, proceso), group in groupby(records, key=lambda r: (r.maquina, r.proceso)):
                group_list = list(group)

                # Sub-agrupar por mes/año
                def get_month_key(r):
                    dt = pd.to_datetime(r.fecha)
                    return (dt.year, dt.month)

                for month_key, month_group in groupby(group_list, key=get_month_key):
                    month_records = list(month_group)
                    
                    # Buscar en DB el último registro del mes ANTERIOR para este grupo
                    # para poder caluclar el delta del primer turno del mes
                    first = month_records[0]
                    dt_first = pd.to_datetime(first.fecha)
                    
                    prev_month_record = self.db.historico.find_first(
                        where={
                            "maquina": maquina,
                            "proceso": proceso,
                            "OR": [
                                {"fecha": {"lt": first.fecha}},
                                {"AND": [{"fecha": first.fecha}, {"turno": {"lt": first.turno}}]}
                            ]
                        },
                        order=[{"fecha": "desc"}, {"turno": "desc"}]
                    )
                    
                    # Verificar que el registro anterior sea del MES anterior (no del mismo)
                    if prev_month_record:
                        dt_prev = pd.to_datetime(prev_month_record.fecha)
                        if dt_prev.year == dt_first.year and dt_prev.month == dt_first.month:
                            prev_month_record = None  # mismo mes, no reiniciar
                    
                    # REGLA: El T1 del último día del mes es el acumulado total del mes.
                    # El T2 del último día se toma como T0 del mes siguiente (no forma parte de este mes).
                    # Por eso, solo procesamos T0 y T1 del mes.
                    # Filtramos T2 del último día del mes (estos NO se cuentan en el mes actual)
                    last_day = max(pd.to_datetime(r.fecha).day for r in month_records)
                    month_records_filtered = [
                        r for r in month_records
                        if not (pd.to_datetime(r.fecha).day == last_day and r.turno == 2)
                    ]

                    prev = prev_month_record

                    for idx, item in enumerate(month_records_filtered):
                        # Obtener los valores acumulados del registro actual
                        curr_vals = {f: (float(getattr(item, f)) if getattr(item, f) is not None else 0.0) for f in FIELDS}
                        prev_vals = {f: (float(getattr(prev, f)) if prev and getattr(prev, f) is not None else None) for f in FIELDS}

                        # Detectar si el siguiente registro (mismo mes) tiene valores MENORES para mts_std
                        # Si es así, el actual es un pico erróneo
                        next_item = None
                        if idx + 1 < len(month_records_filtered):
                            next_item = month_records_filtered[idx + 1]
                        next_vals = {f: (float(getattr(next_item, f)) if next_item and getattr(next_item, f) is not None else None) for f in FIELDS}

                        # Detección de anomalía de duplicidad (>600k en IMPRESION)
                        # Solo usamos los valores escalados para calcular el delta.
                        # NUNCA reescribimos los valores acumulados originales en la BD.
                        curr_mts_eff = curr_vals["prod_metros"]
                        prev_mts_eff = prev_vals["prod_metros"]
                        temp_delta_mts = _calc_delta(curr_mts_eff, prev_mts_eff)

                        scale = 1.0
                        if item.proceso == "IMPRESION" and temp_delta_mts > 600000:
                            print(f"  [DUP] Maquina {item.maquina} ({item.fecha} T{item.turno}): delta {temp_delta_mts:.0f} mts. Usando / 2 solo para delta.")
                            scale = 2.0

                        # Valores efectivos (escalados en memoria, jamás se guardan en la BD)
                        curr_eff = {f: v / scale for f, v in curr_vals.items()}

                        # El T0 del día 1 del mes es el T2 del último día del mes anterior (carryover).
                        # Para este registro, el "drop" hacia el siguiente turno es ESPERADO (el mes nuevo
                        # empieza a acumular desde ese baseline). NO aplicar detección de picos en este caso.
                        dt_item = pd.to_datetime(item.fecha)
                        is_carryover_t0 = (dt_item.day == 1 and item.turno == 0)

                        if is_carryover_t0:
                            # T0 día 1: el agente ya viene reiniciado desde el xlsx (ej. 120,000 mts).
                            # El acumulado de este turno ES su propio delta (no hay anterior en el mes nuevo).
                            # Tratamos prev como None → delta = curr_val directamente.
                            prev_vals = {f: None for f in FIELDS}

                        deltas = {}
                        prev_eff_next = {}  # valores "prev" para el siguiente turno, por campo

                        for cum_field, delta_field in FIELDS.items():
                            c_val = curr_eff[cum_field]
                            p_val = prev_vals[cum_field]
                            n_val = next_vals.get(cum_field)

                            # Pico detectado: el valor actual es mayor al siguiente dentro del mismo mes.
                            is_peak = (n_val is not None and c_val > n_val)

                            if is_peak:
                                # Delta = 0 para este turno, y no avanzamos el prev para este campo
                                deltas[delta_field] = 0.0
                                prev_eff_next[cum_field] = p_val if p_val is not None else 0.0
                            else:
                                deltas[delta_field] = _calc_delta(c_val, p_val)
                                prev_eff_next[cum_field] = c_val

                        self.db.historico.update(where={"id": item.id}, data=deltas)

                        # Avanzar prev con los valores efectivos calculados (en memoria, sin tocar DB)
                        class _PrevProxy:
                            pass
                        prev_proxy = _PrevProxy()
                        for f, v in prev_eff_next.items():
                            setattr(prev_proxy, f, v)
                        prev = prev_proxy

            print("Cálculo de deltas finalizado.")
        finally:
            self.db.disconnect()


def main():
    parser = argparse.ArgumentParser(description="Procesa reportes descargados a la BD SQLite.")
    parser.add_argument("--fecha", required=True, help="Fecha para el reporte (formato YYYY-MM-DD)")
    parser.add_argument("--turno", type=int, choices=[1, 2], required=True, help="Turno (1 o 2)")
    
    args = parser.parse_args()
    
    processor = ProcessReports()
    processor.procesar(args.fecha, args.turno)

if __name__ == "__main__":
    main()
