import streamlit as st
import pandas as pd
import io
from core.engine import build_plan_base, generar_asignacion_dummy, aplicar_y_clasificar_plan

st.set_page_config(page_title="Plan Tejido v5 PRO", layout="wide")

def machine(x):
    try:
        return str(int(float(x))).zfill(4)
    except:
        return str(x).strip()

def parse_date(x):
    return pd.to_datetime(x, errors="coerce").normalize()

st.title("🧵 Plan Tejido v5 PRO (Estable & Optimizado)")

uploaded = st.file_uploader("Sube archivo Excel", type=["xlsx"])

if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        
        # Ingesta flexible de datos
        estado = pd.read_excel(xls, "ESTADO_MAQUINA", header=1)
        demanda = pd.read_excel(xls, "DEMANDA", header=1)
        params = pd.read_excel(xls, "PARAMETROS", header=1)
        
        f_ini = parse_date(params.loc[params["Campo"]=="Fecha_inicio_plan","Valor"].values[0])
        f_fin = parse_date(params.loc[params["Campo"]=="Fecha_fin_plan","Valor"].values[0])
        fechas = pd.date_range(f_ini, f_fin)
        
        maquinas = estado["MAQUINA"].apply(machine).unique().tolist()
        
        # --- PROCESAMIENTO MEDIANTE EL MOTOR ---
        plan = build_plan_base(estado, fechas)
        asign = generar_asignacion_dummy(demanda, fechas, maquinas)
        plan = aplicar_y_clasificar_plan(plan, asign)
        
        # KPIs Diarios
        maquinas_dia = (
            plan.groupby("FECHA")["ACTIVA_DIA"]
            .sum()
            .reset_index(name="MAQS_ACTIVAS")
        )
        
        # --- RENDERIZADO DE INTERFAZ ---
        st.subheader("📋 Plan Detallado")
        st.dataframe(
            plan.sort_values(["MAQUINA", "FECHA"]),
            use_container_width=True,
            height=450
        )
        
        st.subheader("📊 Máquinas Activas (REAL)")
        st.dataframe(maquinas_dia)
        
        # Tabla Pivotizada (Formato Grid Excel)
        pivot = plan.pivot_table(
            index="MAQUINA",
            columns="FECHA",
            values="PLAN_NUEVO_LBS",
            aggfunc="sum",
            fill_value=0
        )
        st.subheader("📊 Vista tipo Excel")
        st.dataframe(pivot)
        
        # --- DESCARGA CONSOLIDADA ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            plan.to_excel(writer, index=False, sheet_name="PLAN_DIARIO")
            maquinas_dia.to_excel(writer, index=False, sheet_name="MAQ_DIA")
            pivot.to_excel(writer, sheet_name="PIVOT")
            
        st.download_button(
            "📥 Descargar Excel",
            data=output.getvalue(),
            file_name="PLAN_TEJIDO_PRO_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"🚨 Error de procesamiento en la estructura de manufactura: {str(e)}")
