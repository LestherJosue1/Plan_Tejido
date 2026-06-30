import pandas as pd
import numpy as np

def build_plan_base(estado_maquina: pd.DataFrame, fechas: pd.DatetimeIndex) -> pd.DataFrame:
    """Genera la matriz base Máquina x Fecha de forma eficiente protegiendo columnas faltantes."""
    # Asegurar limpieza de identificadores de máquina
    estado_maquina["MAQUINA"] = estado_maquina["MAQUINA"].astype(str).str.strip()
    
    # Validar si las columnas existen; si vienen como 'ESTILO' o 'LOTE' se adaptan automáticamente
    for col_req in ["ESTILO_REAL", "LOTE_HILO"]:
        if col_req not in estado_maquina.columns:
            alt_name = "ESTILO" if col_req == "ESTILO_REAL" else "LOTE"
            if alt_name in estado_maquina.columns:
                estado_maquina[col_req] = estado_maquina[alt_name]
            else:
                estado_maquina[col_req] = "" # Columna vacía de respaldo si no existe ninguna
                
    # Crear DataFrame único de máquinas de manera segura
    maquinas_df = estado_maquina[["MAQUINA", "ESTILO_REAL", "LOTE_HILO"]].drop_duplicates()
    fechas_df = pd.DataFrame({"FECHA": fechas})
    
    # Cross join vectorizado (Producto cartesiano)
    base_df = maquinas_df.merge(fechas_df, how="cross")
    
    # Inicializar columnas del plan estructural
    base_df["PLAN_ANTERIOR_ESTILO"] = np.where(base_df["FECHA"] == fechas[0], base_df["ESTILO_REAL"].astype(str).str.strip(), "")
    base_df["PLAN_ANTERIOR_LOTE"] = np.where(base_df["FECHA"] == fechas[0], base_df["LOTE_HILO"].astype(str).str.strip(), "")
    
    base_df["ESTILO_NUEVO"] = ""
    base_df["LOTE_NUEVO"] = ""
    base_df["PLAN_NUEVO_LBS"] = 0.0
    
    return base_df.drop(columns=["ESTILO_REAL", "LOTE_HILO"])


def generar_asignacion_avanzada(demanda: pd.DataFrame, fechas: pd.DatetimeIndex, maquinas: list) -> pd.DataFrame:
    """Motor de asignación basado en prioridades y capacidades dinámicas sin límites fijos."""
    asignaciones = []
    num_maquinas = len(maquinas)
    if num_maquinas == 0 or len(fechas) == 0:
        return pd.DataFrame(columns=["MAQUINA", "FECHA", "LBS", "ESTILO", "LOTE"])
    
    idx_maq = 0
    demanda_sorted = demanda.sort_values(by=["LBS_PENDIENTES"], ascending=False)
    
    for _, d in demanda_sorted.iterrows():
        lbs_pendientes = float(d.get("LBS_PENDIENTES", 0))
        estilo = str(d.get("ESTILO_OPTIMO", "")).strip()
        lote = str(d.get("LOTE_HILO", "")).strip()
        
        for f in fechas:
            pasos = 0
            while lbs_pendientes > 0 and pasos < num_maquinas:
                maq = maquinas[idx_maq]
                prod_dia = min(500.0, lbs_pendientes)
                
                asignaciones.append({
                    "MAQUINA": maq,
                    "FECHA": f,
                    "LBS": prod_dia,
                    "ESTILO": estilo,
                    "LOTE": lote
                })
                
                lbs_pendientes -= prod_dia
                idx_maq = (idx_maq + 1) % num_maquinas  # Balanceo circular de carga
                pasos += 1
                
            if lbs_pendientes <= 0:
                break
                
    return pd.DataFrame(asignaciones) if asignaciones else pd.DataFrame(columns=["MAQUINA", "FECHA", "LBS", "ESTILO", "LOTE"])


def integrar_y_clasificar_plan(plan_df: pd.DataFrame, asignaciones: pd.DataFrame) -> pd.DataFrame:
    """Consolida asignaciones y calcula continuidad/setups de forma 100% vectorizada."""
    if not asignaciones.empty:
        asig_grouped = asignaciones.groupby(["MAQUINA", "FECHA"], as_index=False).agg({
            "LBS": "sum",
            "ESTILO": "first",
            "LOTE": "first"
        })
        
        plan_df = plan_df.merge(asig_grouped, on=["MAQUINA", "FECHA"], how="left")
        plan_df["PLAN_NUEVO_LBS"] = plan_df["LBS"].fillna(0.0)
        plan_df["ESTILO_NUEVO"] = plan_df["ESTILO"].fillna("").str.strip()
        plan_df["LOTE_NUEVO"] = plan_df["LOTE"].fillna("").str.strip()
        plan_df.drop(columns=["LBS", "ESTILO", "LOTE"], inplace=True)

    plan_df = plan_df.sort_values(["MAQUINA", "FECHA"]).reset_index(drop=True)
    
    # Lógica con .shift() para alta velocidad
    plan_df["ESTILO_ACTIVO"] = np.where(plan_df["PLAN_NUEVO_LBS"] > 0, plan_df["ESTILO_NUEVO"], plan_df["PLAN_ANTERIOR_ESTILO"])
    plan_df["ESTILO_ANTERIOR"] = plan_df.groupby("MAQUINA")["ESTILO_ACTIVO"].shift(1).fillna("")
    
    plan_df["CONTINUIDAD"] = (plan_df["ESTILO_ACTIVO"] == plan_df["ESTILO_ANTERIOR"]) & (plan_df["ESTILO_ACTIVO"] != "")
    plan_df["REQUIERE_SETUP"] = (plan_df["ESTILO_ACTIVO"] != plan_df["ESTILO_ANTERIOR"]) & (plan_df["ESTILO_ANTERIOR"] != "") & (plan_df["PLAN_NUEVO_LBS"] > 0)

    condiciones = [(plan_df["PLAN_NUEVO_LBS"] > 0), (plan_df["PLAN_ANTERIOR_ESTILO"].str.strip() != "")]
    elecciones = ["🟢 PRODUCCION", "🔵 CONTINUIDAD"]
    plan_df["TIPO_DIA"] = np.select(condiciones, elecciones, default="⚪ OCIOSO")
    plan_df["ACTIVA_DIA"] = np.where(plan_df["TIPO_DIA"].isin(["🟢 PRODUCCION", "🔵 CONTINUIDAD"]), 1, 0)
    
    plan_df.drop(columns=["ESTILO_ACTIVO", "ESTILO_ANTERIOR"], inplace=True)
    return plan_df
