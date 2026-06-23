import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(page_title="Plan Tejido v4", layout="wide")

# ============================================================
# HELPERS
# ============================================================

def machine(x):
    try:
        return str(int(float(x))).zfill(4)
    except:
        return str(x).strip()

def parse_date(x):
    return pd.to_datetime(x, errors="coerce").normalize()

# ============================================================
# CORE NUEVO MODELO
# ============================================================

def build_plan_base(estado_maquina, fechas):
    records = []

    for _, r in estado_maquina.iterrows():
        maq = machine(r["MAQUINA"])
        estilo = r.get("ESTILO_REAL", "")
        lote = r.get("LOTE_HILO", "")

        for d in fechas:
            records.append({
                "MAQUINA": maq,
                "FECHA": d,
                "PLAN_ANTERIOR_ESTILO": estilo if d == fechas[0] else "",
                "PLAN_ANTERIOR_LOTE": lote if d == fechas[0] else "",
                "ESTILO_NUEVO": "",
                "LOTE_NUEVO": "",
                "PLAN_NUEVO_LBS": 0
            })

    return pd.DataFrame(records)


def generar_asignacion_dummy(demanda, fechas, maquinas):
    # 👉 Simulación simple (puedes reemplazar por tu motor real)
    asignaciones = []

    for _, d in demanda.iterrows():
        lbs = d["LBS_PENDIENTES"]

        for f in fechas:
            for m in maquinas[:2]:  # usa 2 máquinas como ejemplo
                if lbs <= 0:
                    break

                prod = min(500, lbs)

                asignaciones.append({
                    "MAQUINA": m,
                    "FECHA": f,
                    "LBS": prod,
                    "ESTILO": d["ESTILO_OPTIMO"],
                    "LOTE": d["LOTE_HILO"]
                })

                lbs -= prod

    return pd.DataFrame(asignaciones)


def apply_plan_nuevo(plan_df, asignaciones):

    for _, r in asignaciones.iterrows():
        mask = (
            (plan_df["MAQUINA"] == r["MAQUINA"]) &
            (plan_df["FECHA"] == r["FECHA"])
        )

        plan_df.loc[mask, "PLAN_NUEVO_LBS"] += r["LBS"]
        plan_df.loc[mask, "ESTILO_NUEVO"] = r["ESTILO"]
        plan_df.loc[mask, "LOTE_NUEVO"] = r["LOTE"]

    return plan_df


def calcular_continuidad(plan_df):

    plan_df = plan_df.sort_values(["MAQUINA", "FECHA"])
    continuidad = []

    for maq, g in plan_df.groupby("MAQUINA"):

        prev_estilo = None
        prev_lote = None

        for _, r in g.iterrows():

            cont = False

            if r["PLAN_NUEVO_LBS"] > 0:
                if (
                    prev_estilo == r["ESTILO_NUEVO"]
                    and prev_lote == r["LOTE_NUEVO"]
                ):
                    cont = True

                prev_estilo = r["ESTILO_NUEVO"]
                prev_lote = r["LOTE_NUEVO"]

            elif r["PLAN_ANTERIOR_ESTILO"]:
                cont = True
                prev_estilo = r["PLAN_ANTERIOR_ESTILO"]
                prev_lote = r["PLAN_ANTERIOR_LOTE"]

            continuidad.append(cont)

    plan_df["CONTINUIDAD"] = continuidad
    return plan_df


def calcular_activa(plan_df):

    plan_df["ACTIVA_DIA"] = (
        (plan_df["PLAN_NUEVO_LBS"] > 0) |
        (plan_df["CONTINUIDAD"])
    ).astype(int)

    return plan_df


def clasificar_dia(plan_df):

    def tipo(r):
        if r["PLAN_NUEVO_LBS"] > 0:
            return "PRODUCCION"
        elif r["CONTINUIDAD"]:
            return "CONTINUIDAD"
        else:
            return "OCIOSO"

    plan_df["TIPO_DIA"] = plan_df.apply(tipo, axis=1)
    return plan_df


# ============================================================
# STREAMLIT UI
# ============================================================

st.title("🧵 Plan Tejido v4 — con Continuidad Real")

uploaded = st.file_uploader("Sube plantilla", type=["xlsx"])

if uploaded:

    xls = pd.ExcelFile(uploaded)

    estado = pd.read_excel(xls, "ESTADO_MAQUINA", header=1)
    demanda = pd.read_excel(xls, "DEMANDA", header=1)
    params = pd.read_excel(xls, "PARAMETROS", header=1)

    # ===== fechas =====
    f_ini = parse_date(params.loc[params["Campo"]=="Fecha_inicio_plan","Valor"].values[0])
    f_fin = parse_date(params.loc[params["Campo"]=="Fecha_fin_plan","Valor"].values[0])

    fechas = pd.date_range(f_ini, f_fin)

    maquinas = estado["MAQUINA"].apply(machine).unique().tolist()

    # ============================================================
    # RUN MOTOR
    # ============================================================

    plan = build_plan_base(estado, fechas)

    asign = generar_asignacion_dummy(demanda, fechas, maquinas)

    plan = apply_plan_nuevo(plan, asign)
    plan = calcular_continuidad(plan)
    plan = calcular_activa(plan)
    plan = clasificar_dia(plan)

    # ============================================================
    # KPIs
    # ============================================================

    maquinas_dia = (
        plan.groupby("FECHA")["ACTIVA_DIA"]
        .sum()
        .reset_index(name="MAQS_ACTIVAS")
    )

    # ============================================================
    # VISTA
    # ============================================================

    st.subheader("📋 Plan Detallado")

    st.dataframe(
        plan.sort_values(["MAQUINA", "FECHA"]),
        use_container_width=True,
        height=500
    )

    st.subheader("📊 Máquinas Activas por Día")
    st.dataframe(maquinas_dia, use_container_width=True)

    # ============================================================
    # PIVOT tipo tu Excel
    # ============================================================

    pivot = plan.pivot_table(
        index=["MAQUINA"],
        columns="FECHA",
        values="PLAN_NUEVO_LBS",
        aggfunc="sum",
        fill_value=0
    )

    st.subheader("📊 Pivot estilo Excel")
    st.dataframe(pivot)

    # ============================================================
    # DESCARGA
    # ============================================================

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        plan.to_excel(writer, index=False, sheet_name="PLAN_DIARIO")
        maquinas_dia.to_excel(writer, index=False, sheet_name="MAQ_DIA")
        pivot.to_excel(writer, sheet_name="PIVOT")

    st.download_button(
        "📥 Descargar Excel",
        data=output.getvalue(),
        file_name="PLAN_TEJIDO_v4.xlsx"
    )
