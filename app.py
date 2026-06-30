import streamlit as st
import pandas as pd
import io
from core.engine import build_plan_base, generar_asignacion_avanzada, integrar_y_clasificar_plan

st.set_page_config(page_title="Plan Tejido de Avanzada v6", layout="wide")

def clean_machine_series(series: pd.Series) -> pd.Series:
    return series.dropna().astype(float).astype(int).astype(str).str.zfill(4)

def parse_date(x):
    return pd.to_datetime(x, errors="coerce").normalize()

# Capa de optimización de interfaz mediante caché
@st.cache_data(show_spinner="Ejecutando algoritmos de asignación industrial...")
def procesar_planificacion(df_estado, df_demanda, df_params):
    f_ini_val = df_params.loc[df_params["Campo"] == "Fecha_inicio_plan", "Valor"].values[0]
    f_fin_val = df_params.loc[df_params["Campo"] == "Fecha_fin_plan", "Valor"].values[0]
    
    f_ini = parse_date(f_ini_val)
    f_fin = parse_date(f_fin_val)
    fechas = pd.date_range(f_ini, f_fin)
    
    df_estado["MAQUINA"] = clean_machine_series(df_estado["MAQUINA"])
    maquinas_disponibles = df_estado["MAQUINA"].unique().tolist()
    
    # Orquestación del motor matemático
    plan_base = build_plan_base(df_estado, fechas)
    asignaciones = generar_asignacion_avanzada(df_demanda, fechas, maquinas_disponibles)
    plan_final = integrar_y_clasificar_plan(plan_base, asignaciones)
    
    # Compilación de analíticas e indicadores clave
    maquinas_dia = plan_final.groupby("FECHA")["ACTIVA_DIA"].sum().reset_index(name="MAQS_ACTIVAS")
    setups_dia = plan_final.groupby("FECHA")["REQUIERE_SETUP"].sum().reset_index(name="SETUPS_REQUERIDOS")
    kpis_diarios = pd.merge(maquinas_dia, setups_dia, on="FECHA")
    
    pivot_lbs = plan_final.pivot_table(
        index="MAQUINA", columns="FECHA", values="PLAN_NUEVO_LBS", aggfunc="sum", fill_value=0.0
    )
    
    return plan_final, kpis_diarios, pivot_lbs

# --- INTERFAZ GRÁFICA ---
st.title("🧵 Sistema de Planificación Industrial Avanzada — Tejido")
st.markdown("---")

uploaded_file = st.file_uploader("Suba el archivo de control de operaciones de planta (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        pestanas_requeridas = {"ESTADO_MAQUINA", "DEMANDA", "PARAMETROS"}
        if not pestanas_requeridas.issubset(set(xls.sheet_names)):
            st.error(f"Estructura inválida. Faltan pestañas críticas: {pestanas_requeridas}")
            st.stop()
            
        estado = pd.read_excel(xls, "ESTADO_MAQUINA", header=1)
        demanda = pd.read_excel(xls, "DEMANDA", header=1)
        params = pd.read_excel(xls, "PARAMETROS", header=1)
        
        # Procesar datos de manufactura
        plan, kpis, pivot = procesar_planificacion(estado, demanda, params)
        
        # Panel de Métricas de Control Técnico (OEE / Setups)
        st.subheader("⚙️ Métricas de Rendimiento Operativo del Plan")
        m1, m2, m3, m4 = st.columns(4)
        
        total_lbs = plan["PLAN_NUEVO_LBS"].sum()
        total_setups = plan["REQUIERE_SETUP"].sum()
        max_maqs = kpis["MAQS_ACTIVAS"].max()
        eficiencia_uso = (kpis["MAQS_ACTIVAS"].mean() / len(plan["MAQUINA"].unique())) * 100
        
        m1.metric("Libras Planificadas", f"{total_lbs:,.1f} Lbs")
        m2.metric("Eficiencia de Ocupación", f"{eficiencia_uso:.1f} %")
        m3.metric("Picos de Máquinas Activas", f"{max_maqs} Máqs")
        m4.metric("Paradas por Set-up", f"{total_setups} Cambios", delta=int(total_setups), delta_color="inverse")
        
        # Pestañas de Navegación de Datos
        tabs = st.tabs(["📋 Plan de Trabajo Diario", "📊 Matriz de Distribución (Estilo Excel)", "📈 Análisis de Carga por Día"])
        
        with tabs[0]:
            st.dataframe(plan, use_container_width=True, height=400)
        with tabs[1]:
            st.dataframe(pivot, use_container_width=True, height=400)
        with tabs[2]:
            st.markdown("#### Balance de Carga en Planta")
            st.line_chart(kpis.set_index("FECHA"))
            
        # Descarga de Reporte
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            plan.to_excel(writer, index=False, sheet_name="PLAN_DIARIO")
            kpis.to_excel(writer, index=False, sheet_name="METRICAS_DIARIAS")
            pivot.to_excel(writer, sheet_name="PIVOT_MATRIZ")
            
        st.download_button(
            label="📥 Descargar Plan Industrial Validado (.xlsx)",
            data=output.getvalue(),
            file_name="PLAN_TEJIDO_AVANZADO_VALIDADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"🚨 Error crítico en el procesamiento de datos de manufactura: {str(e)}")
        st.info("Verifique que los encabezados y datos del Excel no tengan filas adicionales vacías.")
