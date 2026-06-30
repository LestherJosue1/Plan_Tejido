import pandas as pd
import numpy as np

def build_plan_base(estado_maquina: pd.DataFrame, fechas: pd.DatetimeIndex) -> pd.DataFrame:
    """Genera la matriz base Máquina x Fecha emulando la tolerancia del código original."""
    # Limpieza estándar de identificadores de máquina
    estado_maquina["MAQUINA"] = estado_maquina["MAQUINA"].astype(str).str.strip()
    
    # Tolerancia a variaciones de nombres usando .get() conceptualmente o asignación masiva segura
    estilo_col = "ESTILO_REAL" if "ESTILO_REAL" in estado_maquina.columns else "ESTILO"
    lote_col = "LOTE_HILO" if "LOTE_HILO" in estado_maquina.columns else "LOTE"
    
    records = []
    # Mantenemos una construcción limpia pero estructurada
    for _, r in estado_maquina.iterrows():
        maq = str(r["MAQUINA"]).strip()
        # Tratamiento tolerante idéntico al .get() original
        estilo = str(r.get(estilo_col, "")).strip()
        lote = str(r.get(lote_col, "")).strip()
        
        for i, d in enumerate(fechas):
            records.append({
                "MAQUINA": maq,
                "FECHA": d,
                "PLAN_ANTERIOR_ESTILO": estilo if i == 0 else "",
                "PLAN_ANTERIOR_LOTE": lote if i == 0 else "",
                "ESTILO_NUEVO": "",
                "LOTE_NUEVO": "",
                "PLAN_NUEVO_LBS": 0.0
            })
            
    return pd.DataFrame(records)


def generar_asignacion_dummy(demanda: pd.DataFrame, fechas: pd.DatetimeIndex, maquinas: list) -> pd.DataFrame:
    """Motor de asignación respetando la lógica exacta del código original sin límites duros."""
    asignaciones = []
    # Eliminamos el corte estático [:235] para procesar dinámicamente todo el parque de máquinas disponible
    for _, d in demanda.iterrows():
        lbs = d.get("LBS_PENDIENTES", 0)
        
        for f in fechas:
            for m in maquinas:
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
                
    return pd.DataFrame(asignaciones) if asignaciones else pd.DataFrame(columns=["MAQUINA", "FECHA", "LBS", "ESTILO", "LOTE"])


def aplicar_y_clasificar_plan(plan_df: pd.DataFrame, asignaciones: pd.DataFrame) -> pd.DataFrame:
    """Consolida las asignaciones y calcula estados/continuidades de forma optimizada."""
    if not asignaciones.empty:
        # Agrupación de control por día y telar
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
    
    # --- CÁLCULO DE CONTINUIDAD (Optimizado y alineado al comportamiento v5) ---
    plan_df["ESTILO_ACTIVO"] = np.where(plan_df["PLAN_NUEVO_LBS"] > 0, plan_df["ESTILO_NUEVO"], plan_df["PLAN_ANTERIOR_ESTILO"])
    plan_df["ESTILO_ANTERIOR"] = plan_df.groupby("MAQUINA")["ESTILO_ACTIVO"].shift(1).fillna("")
    
    # Regla de negocio combinada
    plan_df["CONTINUIDAD"] = np.where(
        plan_df["PLAN_NUEVO_LBS"] > 0,
        (plan_df["ESTILO_NUEVO"] == plan_df["ESTILO_ANTERIOR"]) & (plan_df["ESTILO_NUEVO"] != ""),
        (plan_df["PLAN_ANTERIOR_ESTILO"].str.strip() != "")
    )
    
    # Clasificación de tipo de día
    condiciones = [
        (plan_df["PLAN_NUEVO_LBS"] > 0),
        (plan_df["PLAN_ANTERIOR_ESTILO"].astype(str).str.strip() != "")
    ]
    elecciones = ["🟢 PRODUCCION", "🔵 CONTINUIDAD"]
    plan_df["TIPO_DIA"] = np.select(condiciones, elecciones, default="⚪ OCIOSO")
    
    # Indicador binario de máquina activa
    plan_df["ACTIVA_DIA"] = np.where(plan_df["TIPO_DIA"].isin(["🟢 PRODUCCION", "🔵 CONTINUIDAD"]), 1, 0)
    
    plan_df.drop(columns=["ESTILO_ACTIVO", "ESTILO_ANTERIOR"], inplace=True)
    return plan_df
