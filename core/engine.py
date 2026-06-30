import pandas as pd
import numpy as np
import re
import math
from datetime import datetime, date
from fnmatch import fnmatchcase

# ============================================================================
# HELPERS DE NORMALIZACIÓN (Idénticos a tu Colab)
# ============================================================================
def s(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()

def collapse_spaces(t: str) -> str:
    return re.sub(r"\s+", " ", t).strip()

def norm_text(x):
    t = s(x)
    if not t:
        return ""
    return collapse_spaces(t).upper()

def norm_intlike(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    if isinstance(x, (int, np.integer)):
        return str(int(x))
    if isinstance(x, (float, np.floating)):
        if float(x).is_integer():
            return str(int(float(x)))
        return collapse_spaces(str(x)).upper()
    t = collapse_spaces(str(x))
    if re.fullmatch(r"\d+\.0", t):
        return str(int(float(t)))
    return t.upper()

def machine_formatter(x, width=4):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    if isinstance(x, (int, np.integer)):
        t = str(int(x))
    elif isinstance(x, (float, np.floating)) and float(x).is_integer():
        t = str(int(float(x)))
    else:
        t = s(x)
        if re.fullmatch(r"\d+\.0", t):
            t = str(int(float(t)))
    t = t.strip()
    if t.isdigit() and len(t) < width:
        return t.zfill(width)
    return t

def lot_norm(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    if isinstance(x, (int, np.integer)):
        return str(int(x))
    if isinstance(x, (float, np.floating)):
        if float(x).is_integer():
            return str(int(float(x)))
        return collapse_spaces(str(x)).upper()
    t = collapse_spaces(str(x))
    if re.fullmatch(r"\d+\.0", t):
        return str(int(float(t)))
    if "/" in t:
        a, b = t.split("/", 1)
        return a.strip().upper() + "/" + b.strip().upper()
    return t.upper()

def parse_date(x):
    if x is None or (isinstance(x, float) and math.isnan(x)) or s(x) == "":
        return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp, date)):
        return pd.to_datetime(x).normalize()
    return pd.to_datetime(s(x), errors="coerce").normalize()

def split_allowed(val):
    t = norm_text(val)
    if not t:
        return []
    parts = re.split(r"[;,/|]+|\s{2,}", t)
    return [p.strip().upper() for p in parts if p.strip()]

def match_field(rule_val, actual_val, wildcard=False):
    rv = s(rule_val)
    if rv == "" or rv.upper() == "ANY":
        return True
    av = s(actual_val)
    if wildcard:
        return fnmatchcase(av.upper(), rv.upper())
    return av.upper() == rv.upper()

# ============================================================================
# MOTOR PRINCIPAL DE PLANIFICACIÓN
# ============================================================================
def ejecutar_motor_planificacion(df_params, df_estado, df_demanda, df_compat, df_rates, df_reglas, df_restr, df_cal):
    # 1. Parámetros de Control Técnico
    params = {}
    for _, r in df_params.iterrows():
        k = s(r.get("Campo"))
        if not k: continue
        v = r.get("Valor")
        if v is None or (isinstance(v, float) and math.isnan(v)) or s(v) == "":
            v = r.get("Ejemplo")
        params[k] = v

    start_date = parse_date(params.get("Fecha_inicio_plan"))
    end_date = parse_date(params.get("Fecha_fin_plan"))
    hours_day = float(s(params.get("Horas_disponibles_dia") or 24) or 24)
    min_lbs = float(s(params.get("Produccion_min_lbs") or 0) or 0)
    pen_lote = float(s(params.get("Penalizacion_cambio_lote_horas") or 0) or 0)
    pen_estilo = float(s(params.get("Penalizacion_cambio_estilo_horas") or 12) or 12)
    pen_tejido = float(s(params.get("Penalizacion_cambio_tejido_horas") or 0) or 0)
    rate_default = float(s(params.get("Rate_default_lbs_dia") or 1000) or 1000)

    def get_int_param(name, default):
        nums = re.findall(r"\d+", s(params.get(name) or ""))
        return int(nums[0]) if nums else default

    def get_float_param(name, default):
        t = s(params.get(name) or "")
        try: return float(t) if t else default
        except:
            nums = re.findall(r"\d+(\.\d+)?", t)
            return float(nums[0]) if nums else default

    MAX_MAQ_POR_KEY = get_int_param("Max_maquinas_por_key", 30)
    MIN_LBS_NUEVA_MAQ_KEY = get_float_param("Min_lbs_nueva_maquina_por_key", 300.0)
    MIN_HORAS_PROD_BACKFILL = float(s(params.get("Min_horas_produccion_backfill") or 2) or 2)
    TAIL_MAX_LBS = get_float_param("Tail_max_lbs_key", 3500.0)

    # Reglas (SI/NO)
    def rule_on(name, default=True):
        m = df_reglas.loc[df_reglas["Regla"].astype(str).str.strip().str.upper() == name.upper()]
        if m.empty: return default
        v = norm_text(m.iloc[0].get("Valor (SI/NO)"))
        return v in ("SI","S","TRUE","1","YES","Y")

    permitir_setup = rule_on("Permitir_setup_si_hay_demanda", True)
    evitar_peq = rule_on("Evitar_producciones_pequenas", True)
    usar_limite = rule_on("Limite_maquinas_por_dia", False)

    max_machines = None
    max_machines_param = s(params.get("Max_maquinas_dia"))
    if usar_limite and max_machines_param:
        nums = re.findall(r"\d+", max_machines_param)
        if nums: max_machines = max(int(n) for n in nums)

    if pd.isna(start_date) or pd.isna(end_date):
        raise ValueError("PARAMETROS: Fecha_inicio_plan y Fecha_fin_plan son obligatorias.")

    dates = pd.date_range(start_date, end_date, freq="D")
    N = len(dates)

    # 2. Normalizar Demanda
    dem = df_demanda.copy()
    dem["ESTILO"] = dem["ESTILO"].apply(norm_text)
    dem["TITULAR"] = dem["TITULAR"].apply(norm_intlike)
    dem["TEJIDO"] = dem["TEJIDO"].apply(norm_text)
    dem["COLOR"] = dem["COLOR"].apply(norm_text)
    dem["LOTE_HILO"] = dem["LOTE_HILO"].apply(lot_norm)
    dem["FECHA_COMPROMISO"] = dem["FECHA_COMPROMISO"].apply(parse_date) if "FECHA_COMPROMISO" in dem.columns else pd.NaT
    dem["LBS_PENDIENTES"] = pd.to_numeric(dem["LBS_PENDIENTES"], errors="coerce").fillna(0.0).astype(float)
    
    prio_map = {"ALTA":3,"MEDIA":2,"BAJA":1}
    dem["PRIO_NUM"] = dem["PRIORIDAD"].apply(lambda x: prio_map.get(norm_text(x), 2)) if "PRIORIDAD" in dem.columns else 2

    def due_bucket(due):
        if pd.isna(due): return 2
        if due < start_date: return 0
        return 1 if due <= end_date else 2

    dem["_DUEB"] = dem["FECHA_COMPROMISO"].apply(due_bucket)
    dem["_DUED"] = dem["FECHA_COMPROMISO"].fillna(pd.Timestamp.max)

    dem_plan = dem.groupby(["ESTILO","TITULAR","TEJIDO","LOTE_HILO"], dropna=False)["LBS_PENDIENTES"].sum().reset_index().rename(columns={"LBS_PENDIENTES":"LBS_PLAN"})
    key_due = dem.groupby(["ESTILO","TITULAR","TEJIDO","LOTE_HILO"], dropna=False)["FECHA_COMPROMISO"].min().reset_index()
    key_due_map = {(r["ESTILO"], r["TITULAR"], r["TEJIDO"], r["LOTE_HILO"]): r["FECHA_COMPROMISO"] for _, r in key_due.iterrows()}

    key_total = dem_plan.copy()
    key_total["_KEY"] = list(zip(key_total["ESTILO"], key_total["TITULAR"], key_total["TEJIDO"], key_total["LOTE_HILO"]))
    small_keys = set(key_total.loc[key_total["LBS_PLAN"] <= TAIL_MAX_LBS, "_KEY"].tolist())

    # 3. Compatibilidades y Restricciones
    comp = df_compat.copy()
    comp["MAQUINA"] = comp["MAQUINA"].apply(machine_formatter)
    comp["TITULAR"] = comp["TITULAR"].apply(norm_intlike)
    comp["TEJIDO_PERMITIDO"] = comp["TEJIDO_PERMITIDO"].apply(norm_text)
    comp["ACTIVA"] = comp["ACTIVA"].apply(norm_text) if "ACTIVA" in comp.columns else "SI"

    compat_info = {}
    for _, r in comp.iterrows():
        maq = r["MAQUINA"]
        if not maq or r.get("ACTIVA") in ("NO","N","0","FALSE"): continue
        compat_info[maq] = {"allowed": set(split_allowed(r.get("TEJIDO_PERMITIDO"))), "titular": r.get("TITULAR","")}

    if not compat_info:
        raise ValueError("COMPAT_MAQUINA: no hay máquinas activas.")

    restr = df_restr.copy() if not df_restr.empty else pd.DataFrame(columns=["MAQUINA","ESTILO","TITULAR","TEJIDO","LOTE_HILO","PERMITIR","MOTIVO"])
    if not restr.empty:
        for c in ["MAQUINA","ESTILO","TITULAR","TEJIDO","LOTE_HILO","PERMITIR"]:
            if c == "MAQUINA": restr[c] = restr[c].apply(machine_formatter)
            elif c == "LOTE_HILO": restr[c] = restr[c].apply(lot_norm)
            elif c == "TITULAR": restr[c] = restr[c].apply(norm_intlike)
            else: restr[c] = restr[c].apply(norm_text)

    def restriction_blocks(maq, est, tit, tej, lot):
        if restr.empty: return False, ""
        m, e, t, tj, l = machine_formatter(maq), norm_text(est), norm_intlike(tit), norm_text(tej), lot_norm(lot)
        sub = restr[restr["MAQUINA"].fillna("").apply(machine_formatter).isin(["", m])]
        if sub.empty: return False, ""
        best_spec, blocked, motivo = -1, False, ""
        for _, r in sub.iterrows():
            if not (match_field(r.get("MAQUINA",""), m) and match_field(r.get("ESTILO",""), e, wildcard=True) and 
                    match_field(r.get("TITULAR",""), t) and match_field(r.get("TEJIDO",""), tj) and match_field(r.get("LOTE_HILO",""), l, wildcard=True)):
                continue
            spec = sum(1 for c in ["MAQUINA","ESTILO","TITULAR","TEJIDO","LOTE_HILO"] if s(r.get(c)) != "")
            if norm_text(r.get("PERMITIR")) in ("NO","N","0","FALSE") and spec > best_spec:
                best_spec, blocked, motivo = spec, True, s(r.get("MOTIVO"))
        return blocked, motivo

    # 4. Preplan e Historial Heredado (Ocupación)
    preplan_last_date, preplan_free_hours = {}, {}
    preplan_active_day = [set() for _ in range(N)]
    if not df_estado.empty:
        es_pre = df_estado.copy()
        es_pre["MAQUINA"] = es_pre["MAQUINA"].apply(machine_formatter)
        if "FECHA_REF" in es_pre.columns and "HORAS_LIBRES_REF" in es_pre.columns:
            es_pre["FECHA_REF"] = es_pre["FECHA_REF"].apply(parse_date)
            es_pre["HORAS_LIBRES_REF"] = pd.to_numeric(es_pre["HORAS_LIBRES_REF"], errors="coerce")
            for _, r in es_pre.iterrows():
                maq = r.get("MAQUINA", "")
                if not maq or maq not in compat_info: continue
                fd, hf = r.get("FECHA_REF", pd.NaT), r.get("HORAS_LIBRES_REF", np.nan)
                if pd.isna(fd) or pd.isna(hf): continue
                fd = pd.Timestamp(fd).normalize()
                hf = max(0.0, min(float(hf), float(hours_day)))
                preplan_last_date[maq], preplan_free_hours[maq] = fd, hf
                for di, dt in enumerate(dates):
                    if pd.Timestamp(dt).normalize() <= fd: preplan_active_day[di].add(maq)

    # 5. Calendarios de Planta
    cal = df_cal.copy() if not df_cal.empty else pd.DataFrame(columns=["MAQUINA","FECHA_INICIO","FECHA_FIN","TIPO","HORAS_DISPONIBLES"])
    hours_override, blocked_day = {}, set()
    if not cal.empty:
        cal["MAQUINA"] = cal["MAQUINA"].apply(machine_formatter)
        cal["FECHA_INICIO"] = cal["FECHA_INICIO"].apply(parse_date)
        cal["FECHA_FIN"] = cal["FECHA_FIN"].apply(parse_date)
        cal["HORAS_DISPONIBLES"] = pd.to_numeric(cal.get("HORAS_DISPONIBLES"), errors="coerce")
        for _, r in cal.iterrows():
            maq = machine_formatter(r.get("MAQUINA"))
            fi, ff = r.get("FECHA_INICIO"), r.get("FECHA_FIN")
            if pd.isna(fi) or pd.isna(ff) or not maq: continue
            horas = r.get("HORAS_DISPONIBLES")
            if (horas is None or pd.isna(horas)) and norm_text(r.get("TIPO")) in ("MANTENIMIENTO","PARO","STOP","DOWN"):
                horas = 0.0
            elif horas is None or pd.isna(horas): continue
            for di, dt in enumerate(dates):
                if fi <= dt <= ff:
                    if float(horas) <= 0:
                        blocked_day.add((maq, di))
                        hours_override[(maq, di)] = 0.0
                    else: hours_override[(maq, di)] = float(horas)

    # Integración del Preplan Heredado sobre el Calendario Efectivo
    if preplan_last_date:
        date_to_index = {pd.Timestamp(d).normalize(): i for i, d in enumerate(dates)}
        for maq, last_dt in preplan_last_date.items():
            if last_dt not in date_to_index: continue
            last_i = date_to_index[last_dt]
            for di in range(0, last_i):
                blocked_day.add((maq, di))
                hours_override[(maq, di)] = 0.0
            cal_h = float(hours_override.get((maq, last_i), hours_day))
            final_h = min(cal_h, float(preplan_free_hours.get(maq, 0.0)))
            if final_h <= 0:
                blocked_day.add((maq, last_i))
                hours_override[(maq, last_i)] = 0.0
            else: hours_override[(maq, last_i)] = final_h

    def hours_avail(maq, di):
        return float(hours_override.get((maq, di), hours_day))

    # 6. Estructuración Multicapa del Maestro de Eficiencias (RATES)
    rt = df_rates.copy()
    rt["MAQUINA"] = rt["MAQUINA"].apply(machine_formatter)
    rt["ESTILO"] = rt["ESTILO"].apply(norm_text)
    rt["TITULAR"] = rt["TITULAR"].apply(norm_intlike)
    rt["TEJIDO"] = rt["TEJIDO"].apply(norm_text)
    rt["LOTE_HILO"] = rt["LOTE_HILO"].apply(lot_norm)
    rt["TIPO_RATE"] = rt["TIPO_RATE"].apply(norm_text) if "TIPO_RATE" in rt.columns else "RATE"
    rt["FUENTE"] = rt["FUENTE"].apply(norm_text) if "FUENTE" in rt.columns else ""
    rt["RATE_LBS_DIA"] = pd.to_numeric(rt["RATE_LBS_DIA"], errors="coerce")
    rt = rt.dropna(subset=["RATE_LBS_DIA"]).copy()

    def tipo_rank(t):
        if "APROB" in t: return 0
        if "RATE" in t: return 1
        return 2 if "HIST" in t else 3

    rt["_rank"] = rt["TIPO_RATE"].apply(tipo_rank)
    rt = rt.sort_values(by=["_rank"], ascending=True)

    rate_exact, rate_maq_est_tit_tej, rate_maq_est_tej, rate_maq_tej, rate_default_tej = {}, {}, {}, {}, {}
    rate_global = None

    for _, r in rt.iterrows():
        maq, est, tit, tej, lot = r.get("MAQUINA",""), r.get("ESTILO",""), r.get("TITULAR",""), r.get("TEJIDO",""), r.get("LOTE_HILO","")
        rate, fuente = float(r["RATE_LBS_DIA"]), r.get("FUENTE") or r.get("TIPO_RATE") or "RATE"
        if maq and est and tit and tej and lot: rate_exact.setdefault((maq,est,tit,tej,lot), (rate, fuente))
        elif maq and est and tit and tej: rate_maq_est_tit_tej.setdefault((maq,est,tit,tej), (rate, fuente))
        elif maq and est and tej: rate_maq_est_tej.setdefault((maq,est,tej), (rate, fuente))
        elif maq and tej: rate_maq_tej.setdefault((maq,tej), (rate, fuente))
        elif tej: rate_default_tej.setdefault(tej, (rate, fuente))
        elif rate_global is None: rate_global = (rate, fuente)

    def rate_lookup(maq, est, tit, tej, lot):
        maq, est, tit, tej, lot = machine_formatter(maq), norm_text(est), norm_intlike(tit), norm_text(tej), lot_norm(lot)
        if (maq,est,tit,tej,lot) in rate_exact: return rate_exact[(maq,est,tit,tej,lot)][0], rate_exact[(maq,est,tit,tej,lot)][1], "MAQ+EST+TIT+TEJ+LOTE"
        if (maq,est,tit,tej) in rate_maq_est_tit_tej: return rate_maq_est_tit_tej[(maq,est,tit,tej)][0], rate_maq_est_tit_tej[(maq,est,tit,tej)][1], "MAQ+EST+TIT+TEJ"
        if (maq,est,tej) in rate_maq_est_tej: return rate_maq_est_tej[(maq,est,tej)][0], rate_maq_est_tej[(maq,est,tej)][1], "MAQ+EST+TEJ"
        if (maq,tej) in rate_maq_tej: return rate_maq_tej[(maq,tej)][0], rate_maq_tej[(maq,tej)][1], "MAQ+TEJ"
        if tej in rate_default_tej: return rate_default_tej[tej][0], rate_default_tej[tej][1], "DEFAULT_TEJ"
        return (rate_global[0], rate_global[1], "DEFAULT_GLOBAL") if rate_global is not None else (rate_default, "DEFAULT_PARAM", "DEFAULT_PARAM")

    # Inicializar el diccionario de estados actuales de planta
    state = {maq: {"ESTILO":"DISPONIBLE","LOTE_HILO":"DISPONIBLE","TEJIDO":"","TITULAR":info.get("titular",""),"COLOR":""} for maq, info in compat_info.items()}
    if not df_estado.empty:
        es = df_estado.copy()
        es["MAQUINA"] = es["MAQUINA"].apply(machine_formatter)
        for c in ["ESTILO","LOTE_HILO","TITULAR","TEJIDO","COLOR"]:
            if c in es.columns:
                if c == "LOTE_HILO": es[c] = es[c].apply(lot_norm)
                elif c == "TITULAR": es[c] = es[c].apply(norm_intlike)
                else: es[c] = es[c].apply(norm_text)
        for _, r in es.iterrows():
            maq = r["MAQUINA"]
            if not maq or maq not in state: continue
            for c in ["ESTILO","LOTE_HILO","TITULAR","TEJIDO","COLOR"]:
                if c in es.columns and s(r.get(c)) != "": state[maq][c] = r.get(c)

    # 7. Matrices de Trabajo del Motor Secuencial
    schedule = {maq: [[] for _ in range(N)] for maq in compat_info.keys()}
    machines_active_day = [set(preplan_active_day[di]) for di in range(N)]

    def dkey(est, tit, tej, lot): return (norm_text(est), norm_intlike(tit), norm_text(tej), lot_norm(lot))
    def prev_last_key_before_day(maq, i):
        for j in range(i-1, -1, -1):
            if schedule[maq][j]: return schedule[maq][j][-1]
        return state[maq]

    def can_use_day(maq, di):
        if (maq, di) in blocked_day or hours_avail(maq, di) <= 0: return False
        return max_machines is None or maq in machines_active_day[di] or len(machines_active_day[di]) < max_machines

    def machine_can(maq, est, tit, tej, lot):
        info = compat_info[maq]
        if (info.get("titular") and tit and tit != info["titular"]) or (info.get("allowed") and tej and tej not in info["allowed"]): return False
        blocked, _ = restriction_blocks(maq, est, tit, tej, lot)
        return not blocked

    def penalty(prevk, newk):
        prev_est = norm_text(prevk.get("ESTILO"))
        if prev_est in ("","DISPONIBLE"): return 0.0, "INICIO"
        if norm_text(prevk.get("TEJIDO")) and norm_text(prevk.get("TEJIDO")) != norm_text(newk.get("TEJIDO")): return float(pen_tejido), "CAMBIO_TEJIDO"
        if norm_text(prevk.get("ESTILO")) != norm_text(newk.get("ESTILO")): return float(pen_estilo), "CAMBIO_ESTILO"
        if lot_norm(prevk.get("LOTE_HILO")) != lot_norm(newk.get("LOTE_HILO")): return float(pen_lote), "CAMBIO_LOTE"
        return 0.0, "CONTINUIDAD"

    def day_used_hours(maq, di):
        local_hours = hours_avail(maq, di)
        if local_hours <= 0: return 0.0
        used = 0.0
        for seg in schedule[maq][di]:
            used += float(seg["HORAS_SETUP"])
            rate = float(seg["RATE_LBS_DIA"])
            if rate > 0: used += (float(seg["LBS_ASIGNADAS"]) / rate) * local_hours
        return used

    def remaining_hours(maq, di): return max(0.0, hours_avail(maq, di) - day_used_hours(maq, di))
    def capacity_from_hours(rate, local_hours, prod_hours): return float(rate) * (float(prod_hours) / float(local_hours)) if local_hours > 0 else 0.0

    def add_segment(maq, di, keydict, lbs, setup_h, setup_tipo):
        rate, fuente, match = rate_lookup(maq, keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"])
        schedule[maq][di].append({
            "ESTILO": keydict["ESTILO"], "TITULAR": keydict["TITULAR"], "TEJIDO": keydict["TEJIDO"], "LOTE_HILO": keydict["LOTE_HILO"],
            "LBS_ASIGNADAS": float(lbs), "HORAS_SETUP": float(setup_h), "RATE_LBS_DIA": float(rate),
            "FUENTE_RATE": fuente, "RATE_MATCH": match, "SETUP_TIPO": setup_tipo,
        })
        machines_active_day[di].add(maq)

    def earliest_free_day(maq):
        for i in range(N):
            if can_use_day(maq, i) and len(schedule[maq][i]) == 0: return i
        return None

    def assign_primary_block(maq, idx, start_i):
        remaining = float(dem.at[idx, "LBS_PENDIENTES"])
        if remaining <= 0 or start_i is None or start_i >= N: return remaining
        drow = dem.loc[idx]
        if not machine_can(maq, drow["ESTILO"], drow["TITULAR"], drow["TEJIDO"], drow["LOTE_HILO"]): return remaining
        keydict = {"ESTILO":drow["ESTILO"], "TITULAR":drow["TITULAR"], "TEJIDO":drow["TEJIDO"], "LOTE_HILO":drow["LOTE_HILO"]}
        i, first = start_i, True
        while i < N and remaining > 1e-9:
            if not can_use_day(maq, i) or len(schedule[maq][i]) != 0: break
            local_hours = hours_avail(maq, i)
            sh, st = penalty(prev_last_key_before_day(maq, i), keydict)
            if not permitir_setup and sh > 0: break
            horas_netas = max(0.0, local_hours - sh)
            if horas_netas <= 0: break
            rate, _, _ = rate_lookup(maq, keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"])
            lbs = min(remaining, rate * (horas_netas / local_hours))
            if evitar_peq and 0 < lbs < min_lbs and remaining >= min_lbs: break
            if lbs <= 0: break
            add_segment(maq, i, keydict, lbs, sh, ("INICIO" if first else "CONTINUIDAD") if st=="CONTINUIDAD" else st)
            remaining -= lbs
            dem.at[idx, "LBS_PENDIENTES"] = max(0.0, remaining)
            first = False
            i += 1
        return dem.at[idx, "LBS_PENDIENTES"]

    def machines_used_for_key(keydict):
        ms = set()
        for maq in schedule.keys():
            for di in range(N):
                for seg in schedule[maq][di]:
                    if (seg["ESTILO"], seg["TITULAR"], seg["TEJIDO"], seg["LOTE_HILO"]) == (keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"]) and seg["LBS_ASIGNADAS"] > 1e-9:
                        ms.add(maq)
                        break
        return ms

    key_to_idxs = {}
    for idx, row in dem.iterrows():
        key_to_idxs.setdefault(dkey(row["ESTILO"], row["TITULAR"], row["TEJIDO"], row["LOTE_HILO"]), []).append(idx)

    # ------------------------------------------------------------------------
    # FASE 0: Continuidad Heredada Técnica [cite: 30]
    # ------------------------------------------------------------------------
    for maq, st_m in state.items():
        if norm_text(st_m.get("ESTILO")) not in ("", "DISPONIBLE"):
            idxs = key_to_idxs.get(dkey(st_m.get("ESTILO"), st_m.get("TITULAR"), st_m.get("TEJIDO"), st_m.get("LOTE_HILO")), [])
            if idxs: assign_primary_block(maq, idxs[0], earliest_free_day(maq) or 0)

    # ------------------------------------------------------------------------
    # FASE 1: Distribución por Bloques [cite: 30]
    # ------------------------------------------------------------------------
    rem = dem[dem["LBS_PENDIENTES"] > 0].copy()
    group_scores = []
    for g in list(rem.groupby(["TITULAR","TEJIDO"]).groups.keys()):
        sub = rem[(rem["TITULAR"]==g[0]) & (rem["TEJIDO"]==g[1])]
        group_scores.append((g, int(sub["_DUEB"].min()) if not sub.empty else 2, -float(sub["LBS_PENDIENTES"].sum())))
    group_scores.sort(key=lambda x: (x[1], x[2], x[0]))

    for (tit, tej), _, __ in group_scores:
        sub_idx = dem[(dem["TITULAR"]==tit) & (dem["TEJIDO"]==tej) & (dem["LBS_PENDIENTES"]>0)].index.tolist()
        sub_idx.sort(key=lambda i: (dem.at[i, "_DUEB"], dem.at[i, "_DUED"], -dem.at[i, "PRIO_NUM"], -float(dem.at[i, "LBS_PENDIENTES"]), dem.at[i, "ESTILO"], dem.at[i, "LOTE_HILO"]))
        for idx in sub_idx:
            if float(dem.at[idx, "LBS_PENDIENTES"]) <= 0: continue
            drow = dem.loc[idx]
            keyk = (drow["ESTILO"], drow["TITULAR"], drow["TEJIDO"], drow["LOTE_HILO"])
            if keyk in small_keys: continue
            keydict = {"ESTILO":drow["ESTILO"], "TITULAR":drow["TITULAR"], "TEJIDO":drow["TEJIDO"], "LOTE_HILO":drow["LOTE_HILO"]}
            used = list(machines_used_for_key(keydict))
            while float(dem.at[idx, "LBS_PENDIENTES"]) > 1e-9:
                remaining = float(dem.at[idx, "LBS_PENDIENTES"])
                candidates = []
                for maq in schedule.keys():
                    if maq not in used:
                        if len(used) >= MAX_MAQ_POR_KEY or remaining < MIN_LBS_NUEVA_MAQ_KEY: continue
                        si = earliest_free_day(maq)
                        if si is None or not machine_can(maq, keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"]): continue
                        total_cap, i, first = 0.0, si, True
                        while i < N and len(schedule[maq][i]) == 0 and can_use_day(maq, i):
                            sh, _ = penalty(prev_last_key_before_day(maq, i), keydict) if first else (0.0, "CONTINUIDAD")
                            if not permitir_setup and sh > 0: break
                            if max(0.0, hours_avail(maq, i) - sh) <= 0: break
                            rate, _, _ = rate_lookup(maq, keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"])
                            total_cap += rate * (max(0.0, hours_avail(maq, i) - sh) / hours_avail(maq, i))
                            first, i = False, i + 1
                        sh0, _ = penalty(prev_last_key_before_day(maq, si), keydict)
                        candidates.append((maq, 0, total_cap, sh0, si))
                    elif maq in used:
                        si = earliest_free_day(maq)
                        if si is None or not machine_can(maq, keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"]): continue
                        sh0, _ = penalty(prev_last_key_before_day(maq, si), keydict)
                        candidates.append((maq, 1, 999999.0, sh0, si))
                if not candidates: break
                candidates.sort(key=lambda x: (-x[1], -x[2], x[3], x[0]))
                maq_best, _, __, ___, si_best = candidates[0]
                before = float(dem.at[idx, "LBS_PENDIENTES"])
                assign_primary_block(maq_best, idx, si_best)
                if float(dem.at[idx, "LBS_PENDIENTES"]) >= before - 1e-9: break
                if maq_best not in used: used.append(maq_best)

    # ------------------------------------------------------------------------
    # FASE 2: Backfill y Ocupación de Capacidad Remanente [cite: 33]
    # ------------------------------------------------------------------------
    def backfill_one_slot_skip_small(maq, di, max_segments=2):
        if not can_use_day(maq, di) or len(schedule[maq][di]) == 0 or len(schedule[maq][di]) >= max_segments: return False
        free_h = remaining_hours(maq, di)
        if free_h <= 1e-9: return False
        prevk = schedule[maq][di][-1]
        cand_idxs = dem[dem["LBS_PENDIENTES"] > 1e-9].index.tolist()
        cand_idxs.sort(key=lambda i: (dem.at[i, "_DUEB"], dem.at[i, "_DUED"], -dem.at[i, "PRIO_NUM"], -float(dem.at[i, "LBS_PENDIENTES"])))
        local_hours = hours_avail(maq, di)
        for idx in cand_idxs:
            r = dem.loc[idx]
            if (r["ESTILO"], r["TITULAR"], r["TEJIDO"], r["LOTE_HILO"]) in small_keys: continue
            keydict = {"ESTILO":r["ESTILO"], "TITULAR":r["TITULAR"], "TEJIDO":r["TEJIDO"], "LOTE_HILO":r["LOTE_HILO"]}
            if not machine_can(maq, keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"]): continue
            sh2, st2 = penalty(prevk, keydict)
            if (not permitir_setup and sh2 > 0) or free_h < (sh2 + MIN_HORAS_PROD_BACKFILL): continue
            rate2, _, _ = rate_lookup(maq, keydict["ESTILO"], keydict["TITULAR"], keydict["TEJIDO"], keydict["LOTE_HILO"])
            lbs2 = min(float(dem.at[idx, "LBS_PENDIENTES"]), capacity_from_hours(rate2, local_hours, max(0.0, free_h - sh2)))
            if (evitar_peq and 0 < lbs2 < min_lbs and float(dem.at[idx, "LBS_PENDIENTES"]) >= min_lbs) or lbs2 <= 1e-9: continue
            add_segment(maq, di, keydict, lbs2, sh2, st2)
            dem.at[idx, "LBS_PENDIENTES"] = max(0.0, float(dem.at[idx, "LBS_PENDIENTES"]) - lbs2)
            return True
        return False

    for di in range(N):
        for maq in schedule.keys():
            while backfill_one_slot_skip_small(maq, di): pass

    # 8. Transformación a Formato Plano Estructural
    out_rows = []
    for maq, days_list in schedule.items():
        for di, segments in enumerate(days_list):
            f_dt = dates[di]
            h_tot = hours_avail(maq, di)
            if (maq, di) in blocked_day or h_tot <= 0:
                out_rows.append({
                    "MAQUINA": maq, "FECHA": f_dt, "TIPO_DIA": "⚪ OCIOSO", "HORAS_DISPONIBLES": h_tot,
                    "HORAS_PROD": 0.0, "HORAS_SETUP": 0.0, "PLAN_NUEVO_LBS": 0.0, "ESTILO_NUEVO": "",
                    "TITULAR_NUEVO": "", "TEJIDO_NUEVO": "", "LOTE_NUEVO": "", "SETUP_TIPO": "",
                    "RATE_USADO": 0.0, "FUENTE_RATE": "", "MATCH_RATE": ""
                })
                continue
            if not segments:
                p_est = prev_last_key_before_day(maq, di).get("ESTILO","")
                t_dia = "🔵 CONTINUIDAD" if p_est and p_est != "DISPONIBLE" else "⚪ OCIOSO"
                out_rows.append({
                    "MAQUINA": maq, "FECHA": f_dt, "TIPO_DIA": t_dia, "HORAS_DISPONIBLES": h_tot,
                    "HORAS_PROD": 0.0, "HORAS_SETUP": 0.0, "PLAN_NUEVO_LBS": 0.0, "ESTILO_NUEVO": "",
                    "TITULAR_NUEVO": "", "TEJIDO_NUEVO": "", "LOTE_NUEVO": "", "SETUP_TIPO": "",
                    "RATE_USADO": 0.0, "FUENTE_RATE": "", "MATCH_RATE": ""
                })
            else:
                for seg in segments:
                    h_set = float(seg["HORAS_SETUP"])
                    lbs = float(seg["LBS_ASIGNADAS"])
                    rt_u = float(seg["RATE_LBS_DIA"])
                    h_prd = (lbs / rt_u) * h_tot if rt_u > 0 and h_tot > 0 else 0.0
                    out_rows.append({
                        "MAQUINA": maq, "FECHA": f_dt, "TIPO_DIA": "🟢 PRODUCCION", "HORAS_DISPONIBLES": h_tot,
                        "HORAS_PROD": h_prd, "HORAS_SETUP": h_set, "PLAN_NUEVO_LBS": lbs, "ESTILO_NUEVO": seg["ESTILO"],
                        "TITULAR_NUEVO": seg["TITULAR"], "TEJIDO_NUEVO": seg["TEJIDO"], "LOTE_NUEVO": seg["LOTE_HILO"],
                        "SETUP_TIPO": seg["SETUP_TIPO"], "RATE_USADO": rt_u, "FUENTE_RATE": seg["FUENTE_RATE"], "MATCH_RATE": seg["RATE_MATCH"]
                    })

    df_res = pd.DataFrame(out_rows)
    df_res["ACTIVA_DIA"] = np.where(df_res["TIPO_DIA"].isin(["🟢 PRODUCCION", "🔵 CONTINUIDAD"]), 1, 0)
    return df_res
