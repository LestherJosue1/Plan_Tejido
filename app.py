import streamlit as st
import pandas as pd
import io
from core.engine import ejecutar_motor_planificacion

st.set_page_config(page_title="Plan Tejido Colab-Streamlit", layout="wide")

st.title("🧵 Sistema de Planificación — Tejido (Motor Completo Colab)")
st.markdown("Cargue el archivo maestro con las pestañas técnicas requeridas para correr la planificación avanzada.")

uploaded_file = st.file_uploader("Suba el archivo de control de operaciones (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # Validar Pestañas Críticas Obligatorias
        pestanas_criticas = {"PARAMETROS", "ESTADO_MAQUINA", "DEMANDA", "COMPAT_MAQUINA", "RATES", "REGLAS"}
        if not pestanas_criticas.issubset(set(xls.sheet_names)):
            st.error(f"Estructura inválida. Faltan pestañas críticas del sistema: {pestanas_criticas - set(xls.sheet_names)}")
            st.stop()
            
        # Ingesta técnica de datos REPLICANDO EXACTAMENTE EL COMPORTAMIENTO DE TU COLAB
        params_df = pd.read_excel(xls, "PARAMETROS")
        estado_df = pd.read_excel(xls, "ESTADO_MAQUINA")
        demanda_df = pd.read_excel(xls, "DEMANDA")
        compat_df = pd.read_excel(xls, "COMPAT_MAQUINA")
        rates_df = pd.read_excel(xls, "RATES")
        reglas_df = pd.read_excel(xls, "REGLAS")
        
        # Ingesta de opcionales idéntica a Colab
        restr_df = pd.read_excel(xls, "RESTRICCIONES") if "RESTRICCIONES" in xls.sheet_names else pd.DataFrame(columns=["MAQUINA","ESTILO","TITULAR","TEJIDO","LOTE_HILO","PERMITIR","MOTIVO"])
        cal_df = pd.read_excel(xls, "CALENDARIO_MAQUINA") if "CALENDARIO_MAQUINA" in xls.sheet_names else pd.DataFrame(columns=["MAQUINA","FECHA_INICIO","FECHA_FIN","TIPO","HORAS_DISPONIBLES"])

        # Ejecución del Algoritmo del Motor
        with st.spinner("Ejecutando asignaciones y resolviendo restricciones por fases..."):
            plan_final = ejecutar_motor_planificacion(
                params_df, estado_df, demanda_df, compat_df, rates_df, reglas_df, restr_df, cal_df
            )
        
        # --- SECCIÓN DE KPIS MÁXIMOS ---
        st.subheader("⚙️ Métricas Generales Operativas")
        m1, m2, m3, m4 = st.columns(4)
        
        total_lbs = plan_final["PLAN_NUEVO_LBS"].sum()
        maqs_activas_prom = plan_final.groupby("FECHA")["ACTIVA_DIA"].sum().mean()
        total_horas_setup = plan_final["HORAS_SETUP"].sum()
        dias_totales = len(plan_final["FECHA"].unique())
        
        m1.metric("Libras Planificadas", f"{total_lbs:,.1f} Lbs")
        m2.metric("Promedio Máquinas Activas", f"{maqs_activas_prom:.1f} Máqs")
        m3.metric("Horas Invertidas en Setup", f"{total_horas_setup:,.1f} Hrs")
        m4.metric("Días de Cobertura", f"{dias_totales} Días")

        # --- NAVEGACIÓN Y VISUALIZACIONES ---
        tabs = st.tabs(["📋 Plan de Trabajo Diario", "📊 Matriz de Cargas por Máquina", "📈 Resumen de Ocupación"])
        
        with tabs[0]:
            st.dataframe(plan_final, use_container_width=True, height=450)
            
        with tabs[1]:
            pivot_lbs = plan_final.pivot_table(
                index="MAQUINA", columns="FECHA", values="PLAN_NUEVO_LBS", aggfunc="sum", fill_value=0.0
            )
            st.dataframe(pivot_lbs, use_container_width=True, height=450)
            
        with tabs[2]:
            maquinas_dia = plan_final.groupby("FECHA")["ACTIVA_DIA"].sum().reset_index(name="MAQUINAS_ACTIVAS")
            st.line_chart(maquinas_dia.set_index("FECHA"))

        # --- EXPORTADOR EXCEL CON MULTIPESTAÑAS ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            plan_final.to_excel(writer, index=False, sheet_name="PLAN_DIARIO")
            pivot_lbs.to_excel(writer, sheet_name="MATRIZ_PIVOTADA")
            maquinas_dia.to_excel(writer, index=False, sheet_name="KPI_MAQUINAS_DIA")
            
        st.download_button(
            label="📥 Descargar Plan Técnico Validado (.xlsx)",
            data=output.getvalue(),
            file_name="PLAN_TEJIDO_COLAB_PRO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"🚨 Error crítico en el procesamiento del motor de tejido: {str(e)}")
        st.info("Asegúrese de que las columnas tengan los mismos nombres que su cuaderno original de Colab.")
