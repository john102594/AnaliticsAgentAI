import os
import glob
import json
import argparse
import pandas as pd
from prisma import Prisma

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
    def procesar(self, fecha: str, turno: int):
        self.db.connect()
        try:
            fecha_url = fecha.replace("-", "")
            script_dir = os.path.dirname(os.path.abspath(__file__))
            descargas_dir = os.path.join(script_dir, self.config.get("descargas_dir", "descargas"), "reporteturno")
            
            # Buscamos el archivo en descargas
            patron_archivo = f"Novedades_Turno_{turno}_{fecha_url}.xlsx"
            ruta_archivo = os.path.join(descargas_dir, patron_archivo)
            
            if not os.path.exists(ruta_archivo):
                print(f"Archivo no encontrado: {ruta_archivo}")
                return
                
            print(f"Procesando: {ruta_archivo}")
            
            # Ajuste de fecha para DB: Turno 2 se registra con fecha del dia anterior
            orig_fecha = fecha
            orig_turno = turno
            if turno == 2:
                dt_fecha = pd.to_datetime(fecha)
                if dt_fecha.day == 1:
                    turno = 0
                else:
                    fecha = (dt_fecha - pd.Timedelta(days=1)).strftime('%Y-%m-%d')
            
            # Buscar config reporteturno
            reports_cfg = self.config.get("server", {}).get("reports", [])
            reporte_cfg = next((r for r in reports_cfg if r.get("id") == "reporteturno"), {})
            
            # --- Procesar Novedades
            hoja_novedades = reporte_cfg.get("sheet", "IMPRESION")
            try:
                df_nov = pd.read_excel(ruta_archivo, sheet_name=hoja_novedades, header=None, engine="openpyxl")
                inicio_tabla, primera_maquina = self._find_table_start(df_nov)
                fin_tabla = self._find_table_end(df_nov, inicio_tabla)
                
                df_novedades_final = self._process_table_novedades(df_nov, inicio_tabla, fin_tabla, primera_maquina, fecha, turno)
                
                existentes_nov = self.db.novedad.find_many(where={"fecha": fecha, "turno": turno})
                used_nov_ids = set()
                
                count_nov_created = 0
                count_nov_updated = 0
                
                for _, row in df_novedades_final.iterrows():
                    row_dict = row.to_dict()
                    maquina = str(row_dict.pop('Maquina', ''))
                    
                    mapped_data = self._map_novedad_dict(row_dict)
                    create_data = {
                        "fecha": fecha,
                        "turno": turno,
                        "maquina": maquina,
                        "datos": json.dumps(row_dict)
                    }
                    create_data.update(mapped_data)
                    
                    assigned_id = None
                    for ex in existentes_nov:
                        if ex.maquina == maquina and ex.id not in used_nov_ids:
                            assigned_id = ex.id
                            used_nov_ids.add(ex.id)
                            break
                    
                    if assigned_id:
                        update_data = create_data.copy()
                        del update_data["fecha"]
                        del update_data["turno"]
                        del update_data["maquina"]
                        self.db.novedad.update(where={"id": assigned_id}, data=update_data)
                        count_nov_updated += 1
                    else:
                        self.db.novedad.create(data=create_data)
                        count_nov_created += 1
                
                count_nov_deleted = 0
                for ex in existentes_nov:
                    if ex.id not in used_nov_ids:
                        self.db.novedad.delete(where={"id": ex.id})
                        count_nov_deleted += 1
                        
                print(f"Novedades: {count_nov_created} creados, {count_nov_updated} actualizados, {count_nov_deleted} eliminados.")
            except Exception as e:
                print(f"Error procesando novedades: {e}")

            # --- Procesar Historico (SIN CALCULAR DELTAS EN ESTE PASO)
            hojas_historico = reporte_cfg.get("sheets_to_process", ["CORTE", "IMPRESION", "EXLAM", "LAMINACION"])
            
            for sheet_name in hojas_historico:
                try:
                    df_hist = pd.read_excel(ruta_archivo, sheet_name=sheet_name, engine="openpyxl")
                    if df_hist.empty:
                        continue
                        
                    tabla_ext = self._extraer_tabla_historico(df_hist)
                    if tabla_ext is None:
                        continue
                        
                    if sheet_name == "IMPRESION":
                        tabla_ext = tabla_ext.iloc[:, 1:]
                        
                    tabla_limpia = self._limpiar_tabla_historico(tabla_ext, fecha, turno)
                    
                    existentes_hist = self.db.historico.find_many(where={
                        "fecha": fecha, 
                        "turno": turno,
                        "proceso": sheet_name
                    })
                    used_hist_ids = set()
                    
                    count_hist_created = 0
                    count_hist_updated = 0
                    
                    for _, row in tabla_limpia.iterrows():
                        row_dict = row.to_dict()
                        maquina = str(row_dict.pop('Maquina', ''))
                        
                        mapped_data = self._map_historico_dict(row_dict)
                        
                        create_data = {
                            "fecha": fecha,
                            "turno": turno,
                            "maquina": maquina,
                            "proceso": sheet_name,
                            "datos": json.dumps(row_dict)
                        }
                        create_data.update(mapped_data)
                        
                        assigned_id = None
                        for ex in existentes_hist:
                            if ex.maquina == maquina and ex.id not in used_hist_ids:
                                assigned_id = ex.id
                                used_hist_ids.add(ex.id)
                                break
                                
                        if assigned_id:
                            update_data = create_data.copy()
                            del update_data["fecha"]
                            del update_data["turno"]
                            del update_data["maquina"]
                            del update_data["proceso"]
                            self.db.historico.update(where={"id": assigned_id}, data=update_data)
                            count_hist_updated += 1
                        else:
                            self.db.historico.create(data=create_data)
                            count_hist_created += 1
                            
                    count_hist_deleted = 0
                    for ex in existentes_hist:
                        if ex.id not in used_hist_ids:
                            self.db.historico.delete(where={"id": ex.id})
                            count_hist_deleted += 1
                            
                    print(f"Histórico {sheet_name}: {count_hist_created} creados, {count_hist_updated} actualizados, {count_hist_deleted} eliminados.")
                except ValueError:
                    pass
                except Exception as e:
                    print(f"Error procesando historico en {sheet_name}: {e}")

        finally:
            self.db.disconnect()

    def post_procesar_deltas(self, fecha_desde: str):
        """Calcula y actualiza los deltas por turno para todos los registros desde una fecha."""
        self.db.connect()
        try:
            print(f"--- Calculando deltas desde {fecha_desde} ---")
            
            # Obtener todos los registros desde la fecha para recalculas deltas
            todos = self.db.historico.find_many(
                where={"fecha": {"gte": fecha_desde}},
                order=[{"fecha": "asc"}, {"turno": "asc"}]
            )
            
            def _calc_delta(curr, prev):
                if curr is None: return 0.0
                if prev is None: return curr
                return max(curr - prev, 0.0)

            for item in todos:
                # Buscar el registro inmediatamente anterior para MISMA MAQUINA y MISMO PROCESO
                prev = self.db.historico.find_first(
                    where={
                        "maquina": item.maquina,
                        "proceso": item.proceso,
                        "OR": [
                            {"fecha": {"lt": item.fecha}},
                            {"AND": [{"fecha": item.fecha}, {"turno": {"lt": item.turno}}]}
                        ]
                    },
                    order=[{"fecha": "desc"}, {"turno": "desc"}]
                )
                
                # Si es el primer reporte absoluto del mes (dia 1 turno 0), no restamos nada.
                # El Turno 1 del Dia 1 SI debe restar del Turno 0 del Dia 1.
                dt = pd.to_datetime(item.fecha)
                es_inicio_absoluto_mes = dt.day == 1 and item.turno == 0
                
                if es_inicio_absoluto_mes:
                    prev = None

                deltas = {
                    "prod_kg_turno": _calc_delta(item.prod_kg, prev.prod_kg if prev else 0),
                    "prod_metros_turno": _calc_delta(item.prod_metros, prev.prod_metros if prev else 0),
                    "mts_std_turno": _calc_delta(item.mts_std, prev.mts_std if prev else 0),
                    "mts_cargue_turno": _calc_delta(item.mts_cargue, prev.mts_cargue if prev else 0),
                    "desperdicio_turno": _calc_delta(item.desperdicio, prev.desperdicio if prev else 0),
                    "std_desp_turno": _calc_delta(item.std_desp, prev.std_desp if prev else 0),
                    "rechazos_kg_turno": _calc_delta(item.rechazos_kg, prev.rechazos_kg if prev else 0),
                    "rechazos_mts_turno": _calc_delta(item.rechazos_mts, prev.rechazos_mts if prev else 0),
                }
                
                self.db.historico.update(where={"id": item.id}, data=deltas)
            
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
