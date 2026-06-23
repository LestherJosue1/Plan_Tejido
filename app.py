import streamlit as st
import pandas as pd
import io

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(page_title="Plan Tejido v5 PRO", layout="wide")

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
# BASE
# ============================================================

def build_plan_base(estado_maquina, fechas):

    records = []

    for _, r in estado_maquina.iterrows():

        maq = machine(r["MAQUINA"])
        estilo = str(r.get("ESTILO_REAL", "")).strip()
        lote = str(r.get("LOTE_HILO", "")).strip()

        for i, d in enumerate(fechas):

            records.append({
                "MAQUINA": maq,
                "FECHA": d,
                "PLAN_ANTERIOR_ESTILO": estilo if i == 0 else "",
                "PLAN_ANTERIOR_LOTE": lote if i == 0 else "",
                "ESTILO_NUEVO": "",
                "LOTE_NUEVO": "",
                "PLAN_NUEVO_LBS": 0
            })

    return pd.DataFrame(records)

# ============================================================
# MOTOR DEMO
# ============================================================

def generar_asignacion_dummy(demanda, fechas, maquinas):

    asignaciones = []

    for _, d in demanda.iterrows():

        lbs = d.get("LBS_PENDIENTES", 0)

        for f in fechas:
            for m in maquinas[:2]:

                if lbs <= 0:
                    break

                prod = min(500, lbs)

                asignaciones.append({
                    "MAQUINA": m,
                    "FECHA": f,
                    "LBS": prod,
                    "ESTILO": d.get("ESTILO_OPTIMO", ""),
                    "LOTE": d.get("LOTE_HILO", "")
                })

                lbs -= prod

    return pd.DataFrame(asignaciones)

# ============================================================
# APPLY
# ============================================================

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

# ============================================================
# CONTINUIDAD
# ============================================================

def calcular_continuidad(plan_df):

    plan_df = plan_df.sort_values(["MAQUINA", "FECHA"])
    continuidad = []

    for maq, g in plan_df.groupby("MAQUINA"):

        estilo_prev = None

        for _, r in g.iterrows():

            cont = False

            if r["PLAN_NUEVO_LBS"] > 0:
                if estilo_prev == r["ESTILO_NUEVO"]:
                    cont = True
                estilo_prev = r["ESTILO_NUEVO"]

            elif str(r["PLAN_ANTERIOR_ESTILO"]).strip() != "":
                cont = True
                estilo_prev = r["PLAN_ANTERIOR_ESTILO"]

            continuidad.append(cont)

    plan_df["CONTINUIDAD"] = continuidad
    return plan_df

# ============================================================
# ACTIVA
# ============================================================

def calcular_activa(plan_df):

    plan_df["ACTIVA_DIA"] = (
        (plan_df["PLAN_NUEVO_LBS"] > 0) |
        (plan_df["PLAN_ANTERIOR_ESTILO"].astype(str).str.strip() != "")
    ).astype(int)

    return plan_df

# ============================================================
# TIPO DIA (FINAL)
# ============================================================

def clasificar_dia(plan_df):

    def tipo(r):

        if r["PLAN_NUEVO_LBS"] > 0:
            return "🟢 PRODUCCION"

        elif str(r["PLAN_ANTERIOR_ESTILO"]).strip() != "":
            return "🔵 CONTINUIDAD"

        else:
            return "⚪ OCIOSO"

    plan_df["TIPO_DIA"] = plan_df.apply(tipo, axis=1)
    return plan_df

# ============================================================
# APP
# ============================================================

st.title("🧵 Plan Tejido v5 PRO (Estable)")

uploaded = st.file_uploader("Sube archivo Excel", type=["xlsx"])

if uploaded:

    xls = pd.ExcelFile(uploaded)

    estado = pd.read_excel(xls, "ESTADO_MAQUINA", header=1)
    demanda = pd.read_excel(xls, "DEMANDA", header=1)
    params = pd.read_excel(xls, "PARAMETROS", header=1)

    f_ini = parse_date(params.loc[params["Campo"]=="Fecha_inicio_plan","Valor"].values[0])
    f_fin = parse_date(params.loc[params["Campo"]=="Fecha_fin_plan","Valor"].values[0])

    fechas = pd.date_range(f_ini, f_fin)

    maquinas = estado["MAQUINA"].apply(machine).unique().tolist()

    # ============================================================
    # RUN
    # ============================================================

    plan = build_plan_base(estado, fechas)

    asign = generar_asignacion_dummy(demanda, fechas, maquinas)

    plan = apply_plan_nuevo(plan, asign)
    plan = calcular_continuidad(plan)
    plan = calcular_activa(plan)
    plan = clasificar_dia(plan)

    # ============================================================
    # KPI
    # ============================================================

    maquinas_dia = (
        plan.groupby("FECHA")["ACTIVA_DIA"]
        .sum()
        .reset_index(name="MAQS_ACTIVAS")
    )

    # ============================================================
    # VISTA PRINCIPAL (RÁPIDA Y SEGURA)
    # ============================================================

    st.subheader("📋 Plan Detallado")

    st.dataframe(
        plan.sort_values(["MAQUINA", "FECHA"]),
        use_container_width=True,
        height=450
    )

    # ============================================================
    # KPI
    # ============================================================

    st.subheader("📊 Máquinas Activas (REAL)")

    st.dataframe(maquinas_dia)

    # ============================================================
    # PIVOT
    # ============================================================

    pivot = plan.pivot_table(
        index="MAQUINA",
        columns="FECHA",
        values="PLAN_NUEVO_LBS",
        aggfunc="sum",
        fill_value=0
    )

    st.subheader("📊 Vista tipo Excel")

    st.dataframe(pivot)

    # ============================================================
    # EXPORT
    # ============================================================

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        plan.to_excel(writer, index=False, sheet_name="PLAN_DIARIO")
        maquinas_dia.to_excel(writer, index=False, sheet_name="MAQ_DIA")
        pivot.to_excel(writer, sheet_name="PIVOT")

    st.download_button(
        "📥 Descargar Excel",
        data=output.getvalue(),
        file_name="PLAN_TEJIDO_PRO_FINAL.xlsx"
    )
