import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from typing import Dict, Any, Tuple, List, Set

# =====================================================================
# CAPA 1: DATOS E INGESTA (ExcelIngestor)
# =====================================================================
class ExcelIngestor:
    """Clase encargada de la lectura, limpieza y tipado vectorial de datos de planta."""
    
    @staticmethod
    def clean_text_vectorized(series: pd.Series) -> pd.Series:
        return series.astype(str).str.strip().str.replace(r"\s+", " ", regex=True).str.upper()

    @staticmethod
    def clean_machine_vectorized(series: pd.Series, width: int = 4) -> pd.Series:
        # Convierte floats nominales terminados en .0 a enteros nominales limpios
        cleaned = series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        return cleaned.apply(lambda x: x.zfill(width) if x.isdigit() and len(x) < width else x)

    @classmethod
    def process_demanda(cls, df: pd.DataFrame, start_date: datetime, end_date: datetime) -> pd.DataFrame:
        df = df.copy()
        # Detección flexible de la columna de estilo
        estilo_col = "ESTILO_OPTIMO" if "ESTILO_OPTIMO" in df.columns else ("ESTILO" if "ESTILO" in df.columns else df.columns[0])
        df["ESTILO_OPTIMO"] = cls.clean_text_vectorized(df[estilo_col])
        
        df["TEJIDO"] = cls.clean_text_vectorized(df["TEJIDO"])
        df["COLOR"] = cls.clean_text_vectorized(df["COLOR"])
        df["TITULAR"] = cls.clean_text_vectorized(df["TITULAR"])
        df["LOTE_HILO"] = cls.clean_text_vectorized(df["LOTE_HILO"])
        df["LBS_PENDIENTES"] = pd.to_numeric(df["LBS_PENDIENTES"], errors="coerce").fillna(0.0).astype(float)
        
        if "FECHA_COMPROMISO" in df.columns:
            df["FECHA_COMPROMISO"] = pd.to_datetime(df["FECHA_COMPROMISO"], errors="coerce")
        else:
            df["FECHA_COMPROMISO"] = pd.NaT
            
        prio_map = {"ALTA": 3, "MEDIA": 2, "BAJA": 1}
        df["PRIO_NUM"] = cls.clean_text_vectorized(df.get("PRIORIDAD", pd.Series("MEDIA", index=df.index))).map(prio_map).fillna(2)
        
        df["DUE_BUCKET"] = 2
        if not pd.isna(start_date):
            df.loc[df["FECHA_COMPROMISO"] < start_date, "DUE_BUCKET"] = 0
            df.loc[(df["FECHA_COMPROMISO"] >= start_date) & (df["FECHA_COMPROMISO"] <= end_date), "DUE_BUCKET"] = 1
            
        return df

    @classmethod
    def process_compatibilidad(cls, df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
        df = df.copy()
        df["MAQUINA"] = cls.clean_machine_vectorized(df["MAQUINA"])
        df["TEJIDO_PERMITIDO"] = cls.clean_text_vectorized(df["TEJIDO_PERMITIDO"])
        df["TITULAR"] = cls.clean_text_vectorized(df["TITULAR"])
        
        activa_col = cls.clean_text_vectorized(df["ACTIVA"]) if "ACTIVA" in df.columns else pd.Series("SI", index=df.index)
        df = df[activa_col.isin(["SI", "S", "TRUE", "1"])]
        
        compat_map = {}
        for _, row in df.iterrows():
            maq = row["MAQUINA"]
            if not maq or row["MAQUINA"] == "NAN":
                continue
            tejidos = set(re.split(r"[;,/|]+", row["TEJIDO_PERMITIDO"])) if row["TEJIDO_PERMITIDO"] else set()
            compat_map[maq] = {
                "allowed_tejidos": {t.strip() for t in tejidos if t.strip()},
                "titular_fijo": row["TITULAR"] if row["TITULAR"] and row["TITULAR"] != "NAN" else None
            }
        return compat_map

    @classmethod
    def process_restricciones(cls, df: pd.DataFrame) -> Dict[Tuple[str, str], str]:
        """Construye la matriz para el switch dinámico de Estilo Óptimo a Estilo Real."""
        if df.empty:
            return {}
        df = df.copy()
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        # Mapeo posicional o nominal tolerante
        maq_col = "MAQUINA" if "MAQUINA" in df.columns else df.columns[0]
        opt_col = "ESTILO_OPTIMO" if "ESTILO_OPTIMO" in df.columns else ("ESTILO" if "ESTILO" in df.columns else df.columns[1])
        real_col = "ESTILO_REAL" if "ESTILO_REAL" in df.columns else df.columns[2]
        
        df["MAQUINA_CLEAN"] = cls.clean_machine_vectorized(df[maq_col])
        df["OPTIMO_CLEAN"] = cls.clean_text_vectorized(df[opt_col])
        df["REAL_CLEAN"] = cls.clean_text_vectorized(df[real_col])
        
        df_switch = df[df["REAL_CLEAN"].str.len() > 0]
        return dict(zip(zip(df_switch["MAQUINA_CLEAN"], df_switch["OPTIMO_CLEAN"]), df_switch["REAL_CLEAN"]))


# =====================================================================
# CAPA 2: MOTOR DE OPTIMIZACIÓN (AdvancedKnittingEngine)
# =====================================================================
class AdvancedKnittingEngine:
    def __init__(self, config: Dict[str, Any], compat_info: Dict[str, Dict[str, Any]], restricciones_switch: Dict[Tuple[str, str], str]):
        self.config = config
        self.compat_info = compat_info
        self.switch_estilos = restricciones_switch
        self.dates = pd.date_range(config["start_date"], config["end_date"], freq="D")
        self.N = len(self.dates)
        
        self.maquinas = list(compat_info.keys())
        self.schedule = {m: [[] for _ in range(self.N)] for m in self.maquinas}
        self.estilos_procesados_por_maquina = {m: set() for m in self.maquinas}
        self.horas_consumidas_matriz = np.zeros((len(self.maquinas), self.N))
        self.maquinas_index = {m: idx for idx, m in enumerate(self.maquinas)}

    def get_effective_style(self, maquina: str, estilo_optimo: str) -> str:
        """Switch dinámico: Si el óptimo tiene asignado un estilo real en restricciones, lo sustituye."""
        return self.switch_estilos.get((maquina, estilo_optimo), estilo_optimo)

    def calculate_transition_penalty(self, current_key: Dict[str, str], next_key: Dict[str, str]) -> float:
        if current_key["ESTILO"] in ("DISPONIBLE", ""):
            return 0.0
        if current_key["TEJIDO"] != next_key["TEJIDO"]:
            return float(self.config.get("pen_tejido", 0.0))
        if current_key["ESTILO"] != next_key["ESTILO"]:
            return float(self.config.get("pen_estilo", 12.0))
        if current_key["LOTE_HILO"] != next_key["LOTE_HILO"]:
            return float(self.config.get("pen_lote", 0.0))
        return 0.0

    def evaluate_and_assign_vectorized(self, demanda_df: pd.DataFrame) -> Tuple[List[Dict[str, Any]], pd.DataFrame]:
        demanda = demanda_df[demanda_df["LBS_PENDIENTES"] > 0].copy()
        demanda = demanda.sort_values(by=["DUE_BUCKET", "PRIO_NUM", "LBS_PENDIENTES"], ascending=[True, False, False])
        
        asignaciones_totales = []
        
        for idx, row in demanda.iterrows():
            target_lbs = row["LBS_PENDIENTES"]
            opt_style = row["ESTILO_OPTIMO"]
            
            maquinas_feasibles = []
            for m in self.maquinas:
                info = self.compat_info[m]
                if info["titular_fijo"] and row["TITULAR"] != info["titular_fijo"]:
                    continue
                if info["allowed_tejidos"] and row["TEJIDO"] not in info["allowed_tejidos"]:
                    continue
                maquinas_feasibles.append(m)
                
            for m in maquinas_feasibles:
                m_idx = self.maquinas_index[m]
                estilo_efectivo = self.get_effective_style(m, opt_style)
                
                key_req = {
                    "ESTILO": estilo_efectivo, "TEJIDO": row["TEJIDO"], 
                    "LOTE_HILO": row["LOTE_HILO"], "COLOR": row["COLOR"]
                }
                
                for day_idx in range(self.N):
                    cap_dia_max = float(self.config.get("hours_day", 24.0))
                    horas_libres = cap_dia_max - self.horas_consumidas_matriz[m_idx, day_idx]
                    
                    if horas_libres <= 2.0:
                        continue
                        
                    if len(self.schedule[m][day_idx]) > 0:
                        last_seg = self.schedule[m][day_idx][-1]
                        current_state = {"ESTILO": last_seg["ESTILO"], "TEJIDO": last_seg["TEJIDO"], "LOTE_HILO": last_seg["LOTE_HILO"]}
                    else:
                        current_state = {"ESTILO": "", "TEJIDO": "", "LOTE_HILO": ""}
                        
                    penalty_h = self.calculate_transition_penalty(current_state, key_req)
                    horas_utiles = horas_libres - penalty_h
                    
                    if horas_utiles <= 0:
                        continue
                        
                    rate_lbs_dia = float(self.config.get("rate_default", 1000.0))
                    capacidad_libras_disponibles = rate_lbs_dia * (horas_utiles / cap_dia_max)
                    
                    lbs_a_producir = min(target_lbs, capacidad_libras_disponibles)
                    if lbs_a_producir <= 0:
                        continue
                        
                    horas_produccion_nodo = cap_dia_max * (lbs_a_producir / rate_lbs_dia)
                    horas_totales_nodo = horas_produccion_nodo + penalty_h
                    
                    segmento = {
                        "MAQUINA": m, "DIA": self.dates[day_idx].strftime("%Y-%m-%d"),
                        "ESTILO_ORIGINAL": opt_style, "ESTILO_PROCESADO": estilo_efectivo,
                        "TEJIDO": key_req["TEJIDO"], "COLOR": key_req["COLOR"], 
                        "LBS_ASIGNADAS": lbs_a_producir, "HORAS_SETUP": penalty_h, "HORAS_PROD": horas_produccion_nodo
                    }
                    
                    self.schedule[m][day_idx].append(segmento)
                    self.estilos_procesados_por_maquina[m].add(estilo_efectivo)
                    self.horas_consumidas_matriz[m_idx, day_idx] += horas_totales_nodo
                    
                    asignaciones_totales.append(segmento)
                    target_lbs -= lbs_a_producir
                    
                    if target_lbs <= 1e-3:
                        break
                if target_lbs <= 1e-3:
                    break
                    
            demanda.at[idx, "LBS_PENDIENTES"] = target_lbs

        return asignaciones_totales, demanda


# =====================================================================
# CAPA 3: SERVICIOS Y ORQUESTACIÓN (PipelineCoordinator)
# =====================================================================
class PipelineCoordinator:
    
    @staticmethod
    def safe_load_sheet(excel_file, structural_names: list) -> pd.DataFrame:
        xl = pd.ExcelFile(excel_file)
        sheet_to_use = None
        for name in structural_names:
            if name in xl.sheet_names:
                sheet_to_use = name
                break
        if not sheet_to_use:
            raise ValueError(f"Falta hoja crítica en el Excel. Buscadas: {structural_names}")
            
        for try_header in [0, 1, 2]:
            df_test = pd.read_excel(excel_file, sheet_name=sheet_to_use, header=try_header)
            df_test.columns = df_test.columns.astype(str).str.strip().str.upper()
            if any(col in df_test.columns for col in ["ESTILO", "ESTILO_OPTIMO", "TEJIDO", "MAQUINA"]):
                return df_test
        return pd.read_excel(excel_file, sheet_name=sheet_to_use, header=0)

    @classmethod
    def run_pipeline(cls, uploaded_file, parameters: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, int]]:
        xl = pd.ExcelFile(uploaded_file)
        
        # Ingesta Tolerante y Mapeo Dinámico
        df_demanda_raw = cls.safe_load_sheet(uploaded_file, ["DEMANDA", "WIP"])
        
        compat_sheet = "COMPAT_MAQUINA" if "COMPAT_MAQUINA" in xl.sheet_names else ("CUOTA" if "CUOTA" in xl.sheet_names else xl.sheet_names[0])
        df_compat_raw = pd.read_excel(uploaded_file, sheet_name=compat_sheet, header=0)
        df_compat_raw.columns = df_compat_raw.columns.astype(str).str.strip().str.upper()
        if "MAQUINA" not in df_compat_raw.columns and len(xl.sheet_names) > 1:
            df_compat_raw = pd.read_excel(uploaded_file, sheet_name=compat_sheet, header=1)
            df_compat_raw.columns = df_compat_raw.columns.astype(str).str.strip().str.upper()
            
        # Carga opcional de restricciones (Switch de Estilo Real)
        if "RESTRICCIONES" in xl.sheet_names:
            df_restr_raw = pd.read_excel(uploaded_file, sheet_name="RESTRICCIONES", header=0)
        else:
            df_restr_raw = pd.DataFrame()
            
        # Transformaciones de la Capa de Datos
        df_demanda = ExcelIngestor.process_demanda(df_demanda_raw, parameters["start_date"], parameters["end_date"])
        compat_map = ExcelIngestor.process_compatibilidad(df_compat_raw)
        switch_map = ExcelIngestor.process_restricciones(df_restr_raw)
        
        # Ejecución del Proceso de Optimización
        engine = AdvancedKnittingEngine(parameters, compat_map, switch_map)
        asignaciones, df_excedentes = engine.evaluate_and_assign_vectorized(df_demanda)
        
        if asignaciones:
            df_resultados = pd.DataFrame(asignaciones)
        else:
            df_resultados = pd.DataFrame(columns=["MAQUINA", "DIA", "ESTILO_ORIGINAL", "ESTILO_PROCESADO", "TEJIDO", "COLOR", "LBS_ASIGNADAS", "HORAS_SETUP", "HORAS_PROD"])
            
        style_changes_summary = {m: max(0, len(st_set) - 1) for m, st_set in engine.estilos_procesados_por_maquina.items()}
        
        return df_resultados, df_excedentes, style_changes_summary


# =====================================================================
# CAPA 4: PRESENTACIÓN (Streamlit App)
# =====================================================================
def main():
    st.set_page_config(page_title="Advanced Manufacturing Planner - NV2", layout="wide")
    
    st.title("🏭 Planificación Industrial Avanzada & Control de Tejeduría")
    st.subheader("Sistema de Optimización y Asignación de Plan de Tejido (NV2 Core)")
    st.markdown("---")
    
    st.sidebar.header("🎛️ Parámetros del Horizonte Semanal")
    start_date_input = st.sidebar.date_input("Fecha Inicio Plan", datetime(2026, 6, 22))
    end_date_input = st.sidebar.date_input("Fecha Fin Plan", datetime(2026, 6, 28))
    
    hours_day = st.sidebar.slider("Horas Disponibles por Día", 8.0, 24.0, 24.0, step=0.5)
    rate_default = st.sidebar.number_input("Rate Producción Estándar (lbs/día)", min_value=100, max_value=5000, value=1200)
    
    st.sidebar.subheader("⚠️ Penalizaciones por Set-up (Horas)")
    pen_estilo = st.sidebar.number_input("Cambio de Estilo (Horas)", value=12.0, min_value=0.0)
    pen_tejido = st.sidebar.number_input("Cambio de Tejido (Horas)", value=18.0, min_value=0.0)
    pen_lote = st.sidebar.number_input("Cambio de Lote de Hilo (Horas)", value=4.0, min_value=0.0)
    
    config_params = {
        "start_date": pd.Timestamp(start_date_input), "end_date": pd.Timestamp(end_date_input),
        "hours_day": hours_day, "rate_default": rate_default,
        "pen_estilo": pen_estilo, "pen_tejido": pen_tejido, "pen_lote": pen_lote
    }
    
    uploaded_file = st.file_uploader("Cargar Archivo Maestro de Planta (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        with st.spinner("Ejecutando asignación optimizada vectorizada..."):
            try:
                df_plan, df_excedentes, summary_changes = PipelineCoordinator.run_pipeline(uploaded_file, config_params)
                
                st.success("✅ ¡Optimización de Plan de Tejido Generada Exitosamente!")
                
                tab1, tab2, tab3 = st.tabs(["📊 Detalle de Lotes Asignados", "📦 Excedentes (Sin Capacidad)", "⚙️ Indicadores de Maquinaria"])
                
                with tab1:
                    st.dataframe(df_plan, use_container_width=True)
                    
                with tab2:
                    st.warning("Libras que no pudieron ser cubiertas por capacidad física o restricciones:")
                    st.dataframe(df_excedentes.query("LBS_PENDIENTES > 0")[["ESTILO_OPTIMO", "TEJIDO", "COLOR", "LBS_PENDIENTES"]], use_container_width=True)
                    
                with tab3:
                    st.subheader("🔄 Control Crítico de Cambios de Estilos Semanales")
                    summary_df = pd.DataFrame(list(summary_changes.items()), columns=["Código Máquina", "Cambios de Estilo Realizados"])
                    
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.dataframe(summary_df, use_container_width=True)
                    with col2:
                        st.bar_chart(data=summary_df, x="Código Máquina", y="Cambios de Estilo Realizados")
                        
            except Exception as e:
                st.error(f"❌ Error Estructural al procesar el archivo o ejecutar el motor: {str(e)}")

if __name__ == "__main__":
    main()
