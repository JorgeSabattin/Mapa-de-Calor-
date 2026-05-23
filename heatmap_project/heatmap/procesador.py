# ===============================================================
# LÓGICA DE PROCESAMIENTO DE SALAS + HEATMAP
# Adaptado de proyecto_salas.py (versión web Django)
# ===============================================================
import io
import re
import unicodedata
import datetime
import pandas as pd
import numpy as np
import xlsxwriter

# ---------------- UTILIDADES ----------------
INVIS = (160, 8203, 65279)

def strip_invis(s):
    s = "" if s is None else str(s)
    for c in INVIS:
        s = s.replace(chr(c), "")
    return s.strip()

def normtxt(s):
    s = strip_invis(s).lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"[._/()–—-]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def split_slash(cell):
    s = strip_invis(cell)
    if not s:
        return []
    return [p.strip() for p in s.split("/") if p.strip()]

def to_hhmm_any(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, (datetime.datetime, datetime.time, pd.Timestamp)):
        return f"{x.hour:02d}{x.minute:02d}"
    if isinstance(x, (float, int)):
        xf = float(x)
        if 0 <= xf <= 1:
            minutes = int(round(xf * 24 * 60))
            h, m = divmod(minutes, 60)
            return f"{h:02d}{m:02d}"
    s = strip_invis(str(x)).replace("h", ":").replace("H", ":").replace(".", ":")
    s = re.sub(r"\s+", "", s)
    m = re.match(r"^(\d{1,2}):?(\d{2})(?::?\d{0,2})?$", s)
    if m:
        return f"{int(m.group(1)):02d}{int(m.group(2)):02d}"
    return None

def hhmm_to_minutes(hhmm: str) -> int:
    return int(hhmm[:2]) * 60 + int(hhmm[2:])

def hhmm_to_hhmm_colon(hhmm: str) -> str:
    return f"{hhmm[:2]}:{hhmm[2:]}"

def natural_key(text):
    text = "" if text is None else str(text)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]

def cap_pretty(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x)

# ---------------- COLORES ----------------
STRONG_COLORS = [
    "#E53935","#1E88E5","#43A047","#FB8C00","#8E24AA",
    "#00897B","#F4511E","#3949AB","#7CB342","#C2185B",
    "#6D4C41","#546E7A","#D81B60","#5E35B1","#039BE5",
    "#FDD835","#00ACC1","#EF5350","#66BB6A","#FFA726",
    "#8D6E63","#00C853","#FF1744","#2979FF","#AA00FF"
]

# ---------------- MÓDULOS ----------------
mod_horarios_data = {
    'Módulo':  [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20],
    'Inicio':  ['08:30','09:25','10:20','11:15','12:10','13:05','14:00','14:55','15:50','16:45','17:40','18:35','19:30','20:25','19:00','19:46','20:40','21:26','22:20','23:06'],
    'Término': ['09:15','10:10','11:05','12:00','12:55','13:50','14:45','15:40','16:35','17:30','18:25','19:20','20:15','21:10','19:45','20:30','21:25','22:10','23:05','23:50']
}
mods_df = pd.DataFrame(mod_horarios_data)

DAYS = ["Lu","Ma","Mi","Ju","Vi","Sa"]

day_alias = {
    "lu":"Lu","lun":"Lu","lunes":"Lu",
    "ma":"Ma","mar":"Ma","martes":"Ma",
    "mi":"Mi","mie":"Mi","miercoles":"Mi","miércoles":"Mi",
    "ju":"Ju","jue":"Ju","jueves":"Ju",
    "vi":"Vi","vie":"Vi","viernes":"Vi",
    "sa":"Sa","sab":"Sa","sabado":"Sa","sábado":"Sa",
}

day_pat = r"(?:\.?\s*(Lu|Lun|Lunes|Ma|Mar|Martes|Mi|Mie|Mi[eé]rcoles|Ju|Jue|Jueves|Vi|Vie|Viernes|Sa|Sab|S[áa]bado))"
h_pat   = r"(\d{1,2})\s*[:hH:]?\s*(\d{2})"
rng_rx  = re.compile(rf"{day_pat}\s*{h_pat}\s*[-–—]\s*{h_pat}", re.IGNORECASE)

def build_mod_ranges():
    mod_ranges = {}
    for _, r in mods_df.iterrows():
        m = int(r["Módulo"])
        ini = to_hhmm_any(r["Inicio"])
        fin = to_hhmm_any(r["Término"])
        mod_ranges[m] = (ini, fin)
    return mod_ranges

def modulo_a_horario(mod_ranges, m):
    ini, fin = mod_ranges[m]
    return f"{hhmm_to_hhmm_colon(ini)} - {hhmm_to_hhmm_colon(fin)}"

def modules_in_range_by_keys(mod_ranges, hhmm_ini, hhmm_fin, keys):
    s = int(hhmm_ini); e = int(hhmm_fin)
    hits = []
    for m,(mi,mf) in mod_ranges.items():
        if m not in keys:
            continue
        if int(mi) >= s and int(mf) <= e:
            hits.append(m)
    return sorted(set(hits))

def choose_best_group(mod_ranges, hhmm_ini, hhmm_fin, diurno_hits, nocturno_hits):
    s_min = hhmm_to_minutes(hhmm_ini)
    e_min = hhmm_to_minutes(hhmm_fin)
    def score(hits):
        if not hits:
            return (float("inf"), [])
        mi = min(hits); mf = max(hits)
        r_ini = hhmm_to_minutes(mod_ranges[mi][0])
        r_fin = hhmm_to_minutes(mod_ranges[mf][1])
        return (abs(r_ini - s_min) + abs(r_fin - e_min), hits)
    sc_d, hd = score(diurno_hits)
    sc_n, hn = score(nocturno_hits)
    if sc_d == sc_n:
        if int(hhmm_ini) >= 1900:
            return hn if hn else hd
        return hd if hd else hn
    return hd if sc_d < sc_n else hn

def horario_horas_to_occ_sets(mod_ranges, texto):
    occ = {d:set() for d in DAYS}
    if texto is None or not str(texto).strip():
        return occ
    s = strip_invis(str(texto)).replace("—","-").replace("–","-")
    matches = list(rng_rx.finditer(s))
    if not matches:
        return occ
    keys_diurno   = [m for m in mod_ranges.keys() if m <= 14]
    keys_nocturno = [m for m in mod_ranges.keys() if m >= 15]
    for m in matches:
        d_raw = m.group(1)
        d_key = day_alias.get(normtxt(d_raw), d_raw)
        ini = f"{int(m.group(2)):02d}{int(m.group(3)):02d}"
        fin = f"{int(m.group(4)):02d}{int(m.group(5)):02d}"
        diurno_hits   = modules_in_range_by_keys(mod_ranges, ini, fin, keys_diurno)
        nocturno_hits = modules_in_range_by_keys(mod_ranges, ini, fin, keys_nocturno)
        hits = choose_best_group(mod_ranges, ini, fin, diurno_hits, nocturno_hits)
        for mm in hits:
            occ[d_key].add(mm)
    return occ

def make_grid(mod_ranges, occ_day_sets):
    all_mods = list(range(1, 21))
    grid = []
    for m in all_mods:
        rec = {"Modulo": f"M{m:02d}", "Horario": modulo_a_horario(mod_ranges, m)}
        for d in DAYS:
            rec[d] = "OCUP" if m in occ_day_sets[d] else "LIBRE"
        grid.append(rec)
    return pd.DataFrame(grid)

def pct_to_color(pct):
    """Verde (0%) → Amarillo (50%) → Rojo (100%)"""
    pct = max(0.0, min(1.0, pct))
    if pct <= 0.5:
        t = pct * 2
        r = int(99 + (255-99)*t)
        g = int(190 + (235-190)*t)
        b = int(123 + (132-123)*t)
    else:
        t = (pct - 0.5) * 2
        r = int(255 + (248-255)*t)
        g = int(235 + (105-235)*t)
        b = int(132 + (107-132)*t)
    return f"#{r:02X}{g:02X}{b:02X}"


# ===============================================================
# FUNCIÓN PRINCIPAL - devuelve datos para web Y bytes del Excel
# ===============================================================
def procesar_archivo(file_obj):
    """
    Recibe un file-like object (el upload de Django).
    Devuelve un dict con:
      - heatmap_data: lista de edificios con datos para el template
      - indice_data:  lista de salas para el índice
      - excel_bytes:  BytesIO con el Excel generado
      - stats:        estadísticas generales
    """
    df = pd.read_excel(file_obj, header=10)

    required = ["Horario","Edificio","Sala","Capacidad Sala"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas: {missing}. Columnas encontradas: {list(df.columns)}")

    mod_ranges = build_mod_ranges()
    all_mods   = list(range(1, 21))

    # Expandir 1 sala por fila
    rows = []
    for _, r in df.iterrows():
        hs = split_slash(r.get("Horario",""))
        es = split_slash(r.get("Edificio",""))
        ss = split_slash(r.get("Sala",""))
        cs = split_slash(r.get("Capacidad Sala",""))
        n  = max(len(hs), len(es), len(ss), len(cs), 0)

        def pick(lst, i):
            if i < len(lst): return lst[i]
            if len(lst) == 1: return lst[0]
            return None

        for i in range(n):
            h  = pick(hs, i)
            e  = pick(es, i)
            s_ = pick(ss, i)
            c  = pick(cs, i)
            if not (strip_invis(h) and strip_invis(e) and strip_invis(s_)):
                continue
            rows.append({
                "SalaKey": f"{strip_invis(e)}-{strip_invis(s_)}",
                "Edificio": strip_invis(e),
                "Sala": strip_invis(s_),
                "Capacidad": pd.to_numeric(strip_invis(c), errors="coerce"),
                "Horario_ocupado_horas": strip_invis(h),
            })

    df_rooms = pd.DataFrame(rows)
    if df_rooms.empty:
        raise ValueError("No se pudo extraer ninguna sala. Revisa columnas Horario/Edificio/Sala/Capacidad Sala.")

    # Unir ocupación por sala
    occ_union = {}
    meta      = {}
    for row in df_rooms.itertuples(index=False):
        sk  = row.SalaKey
        occ = horario_horas_to_occ_sets(mod_ranges, row.Horario_ocupado_horas)
        if sk not in occ_union:
            occ_union[sk] = {d:set() for d in DAYS}
            meta[sk] = {"Edificio": row.Edificio, "Sala": row.Sala, "Capacidad": row.Capacidad}
        for d in DAYS:
            occ_union[sk][d] |= occ[d]

    grids       = {sk: make_grid(mod_ranges, occ_union[sk]) for sk in occ_union}
    sorted_keys = sorted(grids.keys(), key=natural_key)

    TOTAL_SLOTS = len(all_mods) * len(DAYS)  # 120
    occ_count = {}
    occ_pct   = {}
    for sk in occ_union:
        used = sum(len(occ_union[sk][d]) for d in DAYS)
        occ_count[sk] = used
        occ_pct[sk]   = (used / TOTAL_SLOTS) if TOTAL_SLOTS else 0.0

    room_color_map = {sk: STRONG_COLORS[i % len(STRONG_COLORS)] for i, sk in enumerate(sorted_keys)}

    rooms_by_building = {}
    for sk in sorted_keys:
        ed = meta[sk]["Edificio"]
        rooms_by_building.setdefault(ed, []).append(sk)
    buildings_sorted = sorted(rooms_by_building.keys(), key=natural_key)

    # ---- Datos para el template web ----
    heatmap_data = []
    for ed in buildings_sorted:
        salas = rooms_by_building[ed]
        total = len(salas)
        modulos = []
        for m in all_mods:
            dias = []
            for d in DAYS:
                used = sum(1 for sk in salas if m in occ_union[sk][d])
                pct  = (used / total) if total else 0.0
                dias.append({
                    "dia":   d,
                    "used":  used,
                    "total": total,
                    "pct":   round(pct * 100, 1),
                    "color": pct_to_color(pct),
                })
            modulos.append({
                "num":     m,
                "label":   f"M{m:02d}",
                "horario": modulo_a_horario(mod_ranges, m),
                "dias":    dias,
            })
        heatmap_data.append({
            "edificio": ed,
            "total_salas": total,
            "salas": salas,
            "modulos": modulos,
        })

    # ---- Índice ----
    indice_data = []
    for sk in sorted_keys:
        m = meta[sk]
        indice_data.append({
            "sala_key":  sk,
            "edificio":  m["Edificio"],
            "sala":      m["Sala"],
            "capacidad": cap_pretty(m["Capacidad"]),
            "ocupados":  occ_count[sk],
            "total":     TOTAL_SLOTS,
            "pct":       round(occ_pct[sk] * 100, 1),
            "color":     room_color_map[sk],
        })

    # ---- Salas libres por módulo ----
    salas_libres_data = []
    for m in all_mods:
        dias = []
        for d in DAYS:
            libres = []
            for sk in sorted_keys:
                if m not in occ_union[sk][d]:
                    cap  = cap_pretty(meta.get(sk,{}).get("Capacidad", None))
                    used = occ_count.get(sk, 0)
                    pct  = int(round(occ_pct.get(sk, 0.0) * 100))
                    label = f"{sk} ({cap})" if cap else sk
                    libres.append({"key": sk, "label": label, "color": room_color_map[sk]})
            dias.append({"dia": d, "libres": libres})
        salas_libres_data.append({
            "num":     m,
            "label":   f"M{m:02d}",
            "horario": modulo_a_horario(mod_ranges, m),
            "dias":    dias,
        })

    # ---- Generar Excel ----
    excel_buf = io.BytesIO()
    wb = xlsxwriter.Workbook(excel_buf, {'in_memory': True})

    fmt_hdr    = wb.add_format({"bold":True,"align":"center","valign":"vcenter","border":1})
    fmt_center = wb.add_format({"align":"center","valign":"vcenter"})
    fmt_wrap   = wb.add_format({"text_wrap":True,"valign":"top"})
    fmt_libre  = wb.add_format({"bg_color":"#C6EFCE","align":"center","valign":"vcenter"})
    fmt_ocup   = wb.add_format({"bg_color":"#FFC7CE","align":"center","valign":"vcenter"})
    fmt_pct    = wb.add_format({"num_format":"0%","align":"center","valign":"vcenter"})
    room_fmt   = {sk: wb.add_format({"font_color":room_color_map[sk],"bold":True}) for sk in sorted_keys}

    # Hoja 1: Heatmap
    ws_heat = wb.add_worksheet("Heatmap_uso_por_edificio")
    ws_heat.set_column(0,0,10); ws_heat.set_column(1,1,16); ws_heat.set_column(2,7,12)
    r0 = 0
    for ed in buildings_sorted:
        salas = rooms_by_building[ed]
        total = len(salas)
        ws_heat.merge_range(r0,0,r0,7,f"EDIFICIO {ed}  (salas: {total})",fmt_hdr)
        r0 += 1
        ws_heat.write(r0,0,"Modulo",fmt_hdr); ws_heat.write(r0,1,"Horario",fmt_hdr)
        for j,d in enumerate(DAYS,start=2): ws_heat.write(r0,j,d,fmt_hdr)
        r0 += 1
        start_r = r0
        for m in all_mods:
            ws_heat.write(r0,0,f"M{m:02d}",fmt_center)
            ws_heat.write(r0,1,modulo_a_horario(mod_ranges,m),fmt_center)
            for j,d in enumerate(DAYS,start=2):
                used = sum(1 for sk in salas if m in occ_union[sk][d])
                pct  = (used/total) if total else 0.0
                ws_heat.write_number(r0,j,pct,fmt_pct)
            r0 += 1
        ws_heat.conditional_format(start_r,2,r0-1,7,{
            "type":"3_color_scale","min_color":"#63BE7B","mid_color":"#FFEB84","max_color":"#F8696B"})
        r0 += 2

    # Hoja 2: Salas libres
    ws_free = wb.add_worksheet("Salas_libres_por_modulo")
    ws_free.write(0,0,"Modulo",fmt_hdr); ws_free.write(0,1,"Horario",fmt_hdr)
    for j,d in enumerate(DAYS,start=2): ws_free.write(0,j,d,fmt_hdr)
    ws_free.set_column(0,0,8); ws_free.set_column(1,1,16); ws_free.set_column(2,7,30)
    for i,m in enumerate(all_mods,start=1):
        ws_free.write(i,0,f"M{m:02d}",fmt_center)
        ws_free.write(i,1,modulo_a_horario(mod_ranges,m),fmt_center)
        for j,d in enumerate(DAYS,start=2):
            libres = []
            for sk in sorted_keys:
                if m not in occ_union[sk][d]:
                    cap  = cap_pretty(meta.get(sk,{}).get("Capacidad",None))
                    used = occ_count.get(sk,0)
                    pct  = int(round(occ_pct.get(sk,0.0)*100))
                    label = f"{sk} ({cap}) ({used}/{TOTAL_SLOTS}, {pct}%)" if cap else f"{sk} ({used}/{TOTAL_SLOTS}, {pct}%)"
                    libres.append((sk,label))
            if not libres: continue
            pieces = []
            for k,(sk,label) in enumerate(libres):
                if k > 0: pieces.append("\n")
                pieces.append(room_fmt[sk]); pieces.append(label)
            ws_free.write_rich_string(i,j,*pieces,fmt_wrap)
        ws_free.set_row(i,120)

    # Hoja 3: Índice
    ws_idx = wb.add_worksheet("Indice")
    for c,name in enumerate(["SalaKey","Edificio","Sala","Capacidad","Ocupados","TotalSlots","Ocupacion_%"]):
        ws_idx.write(0,c,name,fmt_hdr)
    for i,item in enumerate(indice_data,start=1):
        ws_idx.write(i,0,item["sala_key"])
        ws_idx.write(i,1,item["edificio"])
        ws_idx.write(i,2,item["sala"])
        ws_idx.write(i,3,item["capacidad"])
        ws_idx.write(i,4,item["ocupados"])
        ws_idx.write(i,5,TOTAL_SLOTS)
        ws_idx.write_number(i,6,occ_pct[item["sala_key"]],fmt_pct)
    ws_idx.set_column(0,0,18); ws_idx.set_column(1,2,10); ws_idx.set_column(3,6,12)

    # Hojas por sala
    for sk in sorted_keys:
        g  = grids[sk]
        ws = wb.add_worksheet(sk[:31])
        for c,col in enumerate(g.columns): ws.write(0,c,col,fmt_hdr)
        ws.set_column(0,0,10); ws.set_column(1,1,16); ws.set_column(2,7,30)
        for rr in range(len(g)):
            ws.write(rr+1,0,g.iloc[rr,0],fmt_center)
            ws.write(rr+1,1,g.iloc[rr,1],fmt_center)
            for cc in range(2,8):
                val = g.iloc[rr,cc]
                ws.write(rr+1,cc,val,fmt_libre if val=="LIBRE" else fmt_ocup)
        ws.freeze_panes(1,2)

    wb.close()
    excel_buf.seek(0)

    stats = {
        "total_salas":    len(sorted_keys),
        "total_edificios": len(buildings_sorted),
        "total_registros": len(df_rooms),
    }

    return {
        "heatmap_data":      heatmap_data,
        "salas_libres_data": salas_libres_data,
        "indice_data":       indice_data,
        "excel_bytes":       excel_buf,
        "stats":             stats,
        "days":              DAYS,
    }
