from __future__ import annotations
import io
import zipfile
import tempfile
import os
import streamlit as st

try:
    import rarfile
    import sys
    if sys.platform == "win32":
        rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"
    # En Linux (Streamlit Cloud) usa el unrar del sistema (instalado via packages.txt)
    _RAR_OK = True
except ImportError:
    _RAR_OK = False
import pandas as pd
import numpy as np
import requests
import urllib3
import psycopg2
import psycopg2.extras

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ════════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════════

# Credenciales Supabase
# En producción (Streamlit Cloud) se leen desde st.secrets
def _pg_config() -> dict:
    """
    Credenciales SOLO desde st.secrets (nunca en el código).
    En Streamlit Cloud: Settings → Secrets. En local: .streamlit/secrets.toml
        [supabase]
        password = "..."
        host = "aws-0-us-east-2.pooler.supabase.com"
        port = 5432
        user = "postgres.pnidiixppefxfiqwzvwz"
    """
    try:
        cfg = st.secrets["supabase"]
        pwd = cfg["password"]
    except Exception:
        st.error("Falta la contraseña de Supabase en los *secrets* de la app "
                 "(sección [supabase], clave password).")
        st.stop()
    return dict(
        host     = cfg.get("host", "aws-0-us-east-2.pooler.supabase.com"),
        port     = int(cfg.get("port", 5432)),
        dbname   = "postgres",
        user     = cfg.get("user", "postgres.pnidiixppefxfiqwzvwz"),
        password = pwd,
        sslmode  = "require",
    )

MES_TXT = [
    "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
    "JULIO","AGOSTO","SETIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE",
]

FECHA_INICIO_INICIAL = pd.Timestamp("2024-12-01")

MOTIVOS_SHEET    = "MOTIVOS_RDO"
COSTOS_SHEET     = "COSTO_RDO"
INDICES_SHEET    = "INDICES_RDO"
GENERACION_SHEET = "GENERACION_RDO"

ARCHIVOS_DEMANDA = {
    "Hidro - Despacho (MW).csv":         "HIDRO",
    "Termica - Despacho (MW).csv":       "TERMICA",
    "Rer y No COES - Despacho (MW).csv": "RER",
}
ARCHIVO_HIDRO   = "Hidro - Despacho (MW).csv"
ARCHIVO_RER     = "Rer y No COES - Despacho (MW).csv"
ARCHIVO_TERMICA = "Termica - Despacho (MW).csv"

RDO_LETRAS_GENERACION = list("ABCDEFGH")

HORAS_GENERACION = [
    f"{((i + 1) * 30) // 60 % 24:02d}:{((i + 1) * 30) % 60:02d}"
    for i in range(48)
]

BARRAS_HIDRO = [
    "CARPAPATA","LA JOYA","STACRUZ12","HUASAHUASI","RONCADOR",
    "PURMACANA","NIMPERIAL","PIZARRAS","POECHOS2","POECHOS1",
    "CANCHAYLLO","CHANCAY","RUCUY","RUNATULLOII","RUNATULLOIII",
    "YANAPAMPA","POTRERO","CH MARANON","YARUCAYA","CHHER1",
    "CHANGELI","CHANGELII","CHANGELIII","8AGOSTO","8/08/2026",
    "RENOVANDESH1","CH RENOVANDES","CH MANTA","SANTA ROSA 1",
    "SANTA ROSA 2","TUPURI","CH HUALLIN",
]
BARRAS_EOLICA = [
    "PE TALARA","PE CUPISNIQUE","PQEEOLICOMARCONA",
    "PQEEOLICO3HERMANAS","WAYRAI","HUAMBOS","DUNA",
    "CE PUNTA LOMITASBL1","CE PUNTA LOMITASBL2",
    "PTALOMITASEXPBL1","PTALOMITASEXPBL2","PE SAN JUAN","WAYRAEXP",
]
BARRAS_SOLAR = [
    "MAJES","TACNASOLAR","PANAMERICANASOLAR","MOQUEGUASOLAR",
    "CS RUBI","INTIPAMPA","CS INTIPAMPA EXPANSION",
    "CSEXPANSIONINTIPAMPA","YARUCAYA","CS YARUCAYA","CSF YARUCAYA",
    "CSCLEMESI","CS CARHUAQUERO","EL CARMEN","CS EL CARMEN",
    "CS MATARANI","CS SAN MARTIN","CSSUNNY","CSSUNNYEXP","LAGRINGAV",
    "CSSUNNYEXPANSION","CSCOENERGY",
]
BARRAS_TERMICA_RER = [
    "MAPLE","PARAMONGA","REPARTICION","HUAYCOLORO","TABLAZO",
    "CTB DONA CATALINA","CT CANA BRAVA","CT SAN JACINTO","CTB CALLAO",
    "CT TALLANCA","CTAGROOLMOS","CASAGRANDE","CSCOENERGY","PIAS",
]

# ════════════════════════════════════════════════════════════════════════════════
# URLS DEL PORTAL COES
# ════════════════════════════════════════════════════════════════════════════════

_BASE_PDO = (
    "https://www.coes.org.pe/portal/browser/download?"
    "url=Operaci%C3%B3n%2FPrograma%20de%20Operaci%C3%B3n%2FPrograma%20Diario%2F"
    "{y}%2F{m}_{M}%2FD%C3%ADa%20{d}%2FYUPANA_{y}{m}{d}.zip"
)
_BASE_RDO = (
    "https://www.coes.org.pe/portal/browser/download?"
    "url=Operaci%C3%B3n%2FPrograma%20de%20Operaci%C3%B3n%2FReprograma%20Diario%20Operaci%C3%B3n%2F"
    "{y}%2F{m}_{M}%2FD%C3%ADa%20{d}%2FReprog%20{d}{m}{letra}%2FYUPANA_{d}{m}{letra}.zip"
)
_BASE_MOTIVO = (
    "https://www.coes.org.pe/portal/browser/download?"
    "url=Operaci%C3%B3n%2FPrograma%20de%20Operaci%C3%B3n%2FReprograma%20Diario%20Operaci%C3%B3n%2F"
    "{y}%2F{m}_{M}%2FD%C3%ADa%20{dd}%2FReprog%20{dd}{mm}{L}%2FReprog_{dd}{mm}{L}.xlsx"
)
_BASE_INDICES = (
    "https://www.coes.org.pe/portal/browser/download?"
    "url=Operaci%C3%B3n%2FPrograma%20de%20Operaci%C3%B3n%2FReprograma%20Diario%20Operaci%C3%B3n%2F"
    "{y}%2F{m}_{M}%2FD%C3%ADa%20{d}%2FReprog%20{dd}{mm}{L}%2Findices{y}{mm}{dd}_{L}.xlsx"
)


def _mes(m: int) -> str:
    return MES_TXT[m - 1]

def _url_pdo(fecha):
    y, m, d = fecha.year, f"{fecha.month:02d}", f"{fecha.day:02d}"
    return _BASE_PDO.format(y=y, m=m, M=_mes(fecha.month), d=d)

def _url_rdo(fecha, letra):
    y, m, d = fecha.year, f"{fecha.month:02d}", f"{fecha.day:02d}"
    return _BASE_RDO.format(y=y, m=m, M=_mes(fecha.month), d=d, letra=letra)

def _url_motivo(fecha, letra):
    y, m, d = fecha.year, f"{fecha.month:02d}", f"{fecha.day:02d}"
    return _BASE_MOTIVO.format(y=y, m=m, M=_mes(fecha.month), d=d, dd=d, mm=m, L=letra)

def _url_indices(fecha, letra):
    y, m, d = fecha.year, f"{fecha.month:02d}", f"{fecha.day:02d}"
    return _BASE_INDICES.format(y=y, m=m, M=_mes(fecha.month), d=d, dd=d, mm=m, L=letra)


# ════════════════════════════════════════════════════════════════════════════════
# HTTP + CACHÉ DE ZIPs
# ════════════════════════════════════════════════════════════════════════════════

_zip_cache: dict = {}

def _get(url: str, timeout: int = 60):
    candidatos = list(dict.fromkeys([url, url.replace("SETIEMBRE", "SEPTIEMBRE")]))
    for u in candidatos:
        try:
            r = requests.get(u, timeout=timeout, verify=False)
            if r.status_code == 200 and r.content:
                return r.content
        except requests.RequestException:
            continue
    return None

def _abrir_zip(data: bytes):
    """Abre ZIP o RAR y devuelve dict {nombre: bytes}."""
    if not data:
        return None

    # — ZIP —
    if data[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                return {n: zf.read(n) for n in zf.namelist()}
        except Exception:
            return None

    # — RAR —
    if data[:7] == b"Rar!\x1a\x07\x00" or data[:8] == b"Rar!\x1a\x07\x01\x00":
        if not _RAR_OK:
            return None
        try:
            with tempfile.NamedTemporaryFile(suffix=".rar", delete=False) as tmp:
                tmp.write(data)
                tmp_path = tmp.name
            with rarfile.RarFile(tmp_path) as rf:
                resultado = {n: rf.read(n) for n in rf.namelist()}
            os.unlink(tmp_path)
            return resultado
        except Exception:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
            return None

    return None

def obtener_zip(fecha, tipo, letra=None):
    key = (fecha.date().isoformat(), tipo, letra)
    if key not in _zip_cache:
        url = _url_pdo(fecha) if tipo == "PDO" else _url_rdo(fecha, letra)
        _zip_cache[key] = _abrir_zip(_get(url)) or {}
    return _zip_cache[key]

def _buscar_en_zip(contenido, nombre):
    nl = nombre.lower()
    for path, data in contenido.items():
        p = path.replace("\\", "/")
        if p.lower().endswith("/" + nl) or p.lower() == nl:
            return data
    return None

def _leer_csv(data):
    if not data:
        return None
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(data), sep=",", encoding=enc, engine="python")
        except Exception:
            pass
    return None


# ════════════════════════════════════════════════════════════════════════════════
# SUPABASE — CONEXIÓN Y HELPERS
# ════════════════════════════════════════════════════════════════════════════════

def _pg_con():
    return psycopg2.connect(**_pg_config())


def _max_fecha_pg(tabla: str):
    """SELECT MAX(Fecha) de una tabla en Supabase."""
    try:
        con = _pg_con()
        cur = con.cursor()
        cur.execute(f'SELECT MAX("Fecha") FROM "{tabla}"')
        val = cur.fetchone()[0]
        con.close()
        return pd.Timestamp(val) if val else None
    except Exception:
        return None


_cols_cache: dict = {}

def _columnas_tabla(con, tabla: str) -> list:
    """Columnas que existen hoy en la tabla de Supabase (para insertar solo esas)."""
    if tabla not in _cols_cache:
        cur = con.cursor()
        cur.execute("SELECT column_name, data_type FROM information_schema.columns "
                    "WHERE table_schema = 'public' AND table_name = %s", (tabla,))
        _cols_cache[tabla] = dict(cur.fetchall())
    return _cols_cache[tabla]


def _insertar_df_pg(df: pd.DataFrame, tabla: str, con) -> int:
    """
    Reemplaza por día: borra las filas de las fechas del df y las vuelve a insertar.
    Así el último día guardado siempre queda completo (p. ej. si salió un RDO nuevo)
    y no quedan filas viejas. Todo en una transacción: si algo falla, no se borra nada.
    Solo se insertan las columnas que existen en la tabla (las que Power BI
    reconstruye solas, como Hora, se omiten si ya no están en Supabase).
    Retorna el número de filas insertadas.
    """
    if df is None or df.empty:
        return 0
    df = df.copy()
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.strftime("%Y-%m-%d")
    existentes = _columnas_tabla(con, tabla)
    if existentes:
        df = df[[c for c in df.columns if c in existentes]]
    df = df.astype(object).where(pd.notnull(df), None)    # NaN → NULL

    cols = list(df.columns)
    cols_sql = ", ".join('"' + c + '"' for c in cols)
    sql_ins = f'INSERT INTO "{tabla}" ({cols_sql}) VALUES %s'
    fechas = sorted(df["Fecha"].dropna().unique().tolist())
    cur = con.cursor()
    try:
        if fechas:
            tipo_fecha = "::date[]" if existentes.get("Fecha") == "date" else ""
            cur.execute(f'DELETE FROM "{tabla}" WHERE "Fecha" = ANY(%s{tipo_fecha})', (fechas,))
        rows = [tuple(r) for r in df.itertuples(index=False)]
        psycopg2.extras.execute_values(cur, sql_ins, rows, page_size=2000)
        con.commit()
    except Exception:
        con.rollback()
        raise
    return len(df)


# ════════════════════════════════════════════════════════════════════════════════
# UTILIDADES COMPARTIDAS
# ════════════════════════════════════════════════════════════════════════════════

def fila_sin_primer_valor(df):
    if df is None or df.empty:
        return None
    if df.shape[1] > 1:
        return df.iloc[:, 1:].apply(pd.to_numeric, errors="coerce").sum(axis=1).tolist()
    total = []
    for celda in df.iloc[:, 0].astype(str):
        nums = [float(x) for x in celda.split(",")[1:] if x.strip()]
        total.append(sum(nums))
    return total

def rellenar_hasta_48(lst):
    if not lst:
        return None
    faltan = 48 - len(lst)
    return ([0] * faltan + lst) if faltan > 0 else lst[:48]


# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — DEMANDA
# ════════════════════════════════════════════════════════════════════════════════

def _procesar_csv_de_zip(fecha, tipo, letra, nombre_csv, fuente):
    contenido = obtener_zip(fecha, tipo, letra)
    data = _buscar_en_zip(contenido, nombre_csv)
    if data is None:
        return pd.DataFrame()
    df = _leer_csv(data)
    valores = fila_sin_primer_valor(df)
    if not valores:
        return pd.DataFrame()
    valores_48 = rellenar_hasta_48(valores)
    if valores_48 is None:
        return pd.DataFrame()
    anio    = fecha.year
    fecha_d = str(fecha.date())
    if tipo == "PDO":
        path    = f"PDO/{anio}/{fecha_d}/{nombre_csv}"
        rdo_val = None
    else:
        path    = f"RDO/{anio}/{fecha_d}/{letra}/{nombre_csv}"
        rdo_val = letra          # ← solo la letra (A, B…), sin "RDO " al inicio
    return pd.DataFrame({
        "Año":       anio,
        "Fecha":     fecha_d,
        "TipoSerie": tipo,
        "RDO":       rdo_val,
        "Archivo":   nombre_csv,
        "Periodo":   range(1, 49),
        "MW":        valores_48,
        "Fuente":    fuente,
        "path":      path,
    })


def procesar_demanda_rango(fecha_inicio, fecha_fin, log):
    fechas = pd.date_range(fecha_inicio, fecha_fin, freq="D")
    total  = len(fechas)
    tablas = []
    log.write(f"DEMANDA — {fecha_inicio.date()} → {fecha_fin.date()} ({total} días)")
    prog = st.progress(0)
    for i, fdt in enumerate(fechas, 1):
        prog.progress(i / total)
        subtablas = []
        for nombre_csv, fuente in ARCHIVOS_DEMANDA.items():
            df = _procesar_csv_de_zip(fdt, "PDO", None, nombre_csv, fuente)
            if not df.empty:
                subtablas.append(df)
        for letra in RDO_LETRAS_GENERACION:
            for nombre_csv, fuente in ARCHIVOS_DEMANDA.items():
                df = _procesar_csv_de_zip(fdt, "RDO", letra, nombre_csv, fuente)
                if not df.empty:
                    subtablas.append(df)
        if subtablas:
            tablas.append(pd.concat(subtablas, ignore_index=True))
    prog.empty()
    return pd.concat(tablas, ignore_index=True) if tablas else pd.DataFrame()


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — MOTIVOS RDO
# ════════════════════════════════════════════════════════════════════════════════

def tab2_extraer_hora(df):
    try:
        val = df.iat[6, 1]
        return "" if pd.isna(val) else str(val).strip()
    except Exception:
        return ""


def tab2_extraer_motivo(df):
    import re
    if df is None or df.empty:
        return ""
    s = lambda x: "" if pd.isna(x) else str(x).strip()
    motivo = ""
    try:
        if df.shape[1] >= 3:
            colC = df.iloc[:, 2].astype(str).str.upper()
            idx  = colC[colC.str.contains("MOTIVO", na=False)].index
            if len(idx) > 0:
                fila = idx[0] + 1
                if fila < len(df):
                    for celda in df.iloc[fila, :].map(s).tolist():
                        if "-" in celda and re.search(r"-\s*[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", celda):
                            motivo = celda.strip()
                            break
    except Exception:
        pass
    if not motivo:
        try:
            if df.shape[1] >= 4:
                colD = [s(v) for v in df.iloc[:, 3].tolist()]
                colD = [v for v in colD if v]
                if colD:
                    motivo = colD[-1]
        except Exception:
            pass
    return motivo


def tab2_procesar_motivos_fecha(fecha) -> pd.DataFrame:
    fecha_dt = pd.Timestamp(fecha)
    d = f"{fecha_dt.day:02d}"
    m = f"{fecha_dt.month:02d}"
    registros = []
    for letra in "ABCDEFG":
        url  = _url_motivo(fecha_dt, letra)
        data = _get(url)
        if data is None or data[:4] != b"PK\x03\x04":
            continue
        try:
            df = pd.read_excel(io.BytesIO(data), header=None, engine="openpyxl")
            if df is None or df.empty:
                continue
            registros.append({
                "Fecha":   str(fecha),
                "Hora":    tab2_extraer_hora(df),
                "TipoRDO": f"RDO {letra}",
                "Motivo":  tab2_extraer_motivo(df),
                "Archivo": f"Reprog_{d}{m}{letra}.xlsx",
                "Ruta":    f"Motivos RDO/{fecha_dt.year}/{fecha}/Reprog_{d}{m}{letra}.xlsx",
            })
        except Exception:
            pass
    return pd.DataFrame(registros, columns=["Fecha","Hora","TipoRDO","Motivo","Archivo","Ruta"])


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — COSTO RDO
# ════════════════════════════════════════════════════════════════════════════════

TAB2_NOMBRE_COSTO = "Termica - Costo Operacion por Etapas ($).csv"

def _tab2_leer_costo(data):
    try:
        df = pd.read_csv(io.BytesIO(data))
        return df.apply(pd.to_numeric, errors="coerce")
    except Exception:
        return None

def _tab2_sumar_costo(df_num):
    if df_num is None or df_num.shape[1] <= 1:
        return 0.0
    return float(df_num.iloc[:, 1:].sum(axis=1, skipna=True).sum(skipna=True))

def _costo_del_zip(contenido):
    data = _buscar_en_zip(contenido, TAB2_NOMBRE_COSTO)
    if data is None:
        return None, 0, None
    df_num = _tab2_leer_costo(data)
    if df_num is None:
        return None, 0, None
    return _tab2_sumar_costo(df_num), len(df_num), df_num

def tab2_procesar_costos_fecha(fecha) -> pd.DataFrame:
    fecha_dt = pd.Timestamp(fecha)
    anio     = fecha_dt.year
    zip_pdo  = obtener_zip(fecha_dt, "PDO")
    subtotal_pdo, filas_pdo, df_pdo = _costo_del_zip(zip_pdo)
    if subtotal_pdo is None:
        return pd.DataFrame()
    resultados = [{
        "Fecha": str(fecha), "TipoSerie": "PDO", "RDO": "",
        "Subtotal": subtotal_pdo, "Filas": filas_pdo, "FilasPDO": filas_pdo,
        "DiferenciaFilas": 0, "CostoAdicionalPDO": 0.0, "CostoTotal": subtotal_pdo,
        "Archivo": TAB2_NOMBRE_COSTO,
        "Ruta": f"PDO/{anio}/{fecha}/{TAB2_NOMBRE_COSTO}",
    }]
    for letra in "ABCDEFG":
        zip_rdo = obtener_zip(fecha_dt, "RDO", letra)
        subtotal_rdo, filas_rdo, _ = _costo_del_zip(zip_rdo)
        if subtotal_rdo is None:
            continue
        diferencia      = filas_pdo - filas_rdo
        costo_adicional = _tab2_sumar_costo(df_pdo.head(diferencia)) if diferencia > 0 and df_pdo is not None else 0.0
        resultados.append({
            "Fecha": str(fecha), "TipoSerie": "RDO", "RDO": f"RDO {letra}",
            "Subtotal": subtotal_rdo, "Filas": filas_rdo, "FilasPDO": filas_pdo,
            "DiferenciaFilas": diferencia, "CostoAdicionalPDO": costo_adicional,
            "CostoTotal": subtotal_rdo + costo_adicional,
            "Archivo": TAB2_NOMBRE_COSTO,
            "Ruta": f"RDO/{anio}/{fecha}/{letra}/{TAB2_NOMBRE_COSTO}",
        })
    return pd.DataFrame(resultados)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — INDICES ALFA/BETA/GAMMA
# ════════════════════════════════════════════════════════════════════════════════

def _tab2_encontrar_col(headers, palabra):
    palabra = palabra.lower()
    for i, h in enumerate(headers):
        if h == palabra: return i
    for i, h in enumerate(headers):
        if palabra in h:  return i
    return None

def tab2_procesar_indices_fecha(fecha) -> pd.DataFrame:
    fecha_dt = pd.Timestamp(fecha)
    d = f"{fecha_dt.day:02d}"; m = f"{fecha_dt.month:02d}"
    nombre = ruta = data = None
    for letra in "GFEDCBA":
        url = _url_indices(fecha_dt, letra)
        raw = _get(url)
        if raw and raw[:4] == b"PK\x03\x04":
            nombre = f"indices{fecha_dt.year}{m}{d}_{letra}.xlsx"
            ruta   = f"Indices/{fecha_dt.year}/{fecha}/{nombre}"
            data   = raw
            break
    if data is None:
        return pd.DataFrame()
    try:
        df = pd.read_excel(io.BytesIO(data), header=None, engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()
    headers = ["" if pd.isna(x) else str(x).strip().lower() for x in df.iloc[0, :].tolist()]
    c_alfa  = _tab2_encontrar_col(headers, "alfa")
    c_beta  = _tab2_encontrar_col(headers, "beta")
    c_gamma = _tab2_encontrar_col(headers, "gamma")
    def col_vals(idx):
        if idx is None: return [np.nan] * 48
        vals = df.iloc[1:49, idx].tolist()
        out  = []
        for v in vals:
            if pd.isna(v): out.append(np.nan)
            else:
                try:    out.append(float(v))
                except: out.append(np.nan)
        while len(out) < 48: out.append(np.nan)
        return out[:48]
    alfa = col_vals(c_alfa); beta = col_vals(c_beta); gamma = col_vals(c_gamma)
    horas = []
    h, minuto = 0, 30
    for _ in range(48):
        horas.append(f"{h:02d}:{minuto:02d}")
        minuto += 30
        if minuto == 60: minuto = 0; h = (h + 1) % 24
    return pd.DataFrame([{
        "Fecha": str(fecha), "ArchivoIndices": nombre, "Periodo": i + 1,
        "Hora": horas[i], "ALFA": alfa[i], "BETA": beta[i], "GAMMA": gamma[i], "Ruta": ruta,
    } for i in range(48)])


def tab2_procesar_rango(fecha_inicio, fecha_fin, log):
    fechas = pd.date_range(pd.Timestamp(fecha_inicio), pd.Timestamp(fecha_fin), freq="D")
    total  = len(fechas)
    m_lst, c_lst, i_lst = [], [], []
    log.write(f"MOTIVOS / COSTOS / ÍNDICES — {fecha_inicio.date()} → {fecha_fin.date()} ({total} días)")
    prog = st.progress(0)
    for i, fdt in enumerate(fechas, 1):
        prog.progress(i / total)
        fecha = fdt.strftime("%Y-%m-%d")
        df_m = tab2_procesar_motivos_fecha(fecha)
        if not df_m.empty: m_lst.append(df_m)
        df_c = tab2_procesar_costos_fecha(fecha)
        if not df_c.empty: c_lst.append(df_c)
        df_i = tab2_procesar_indices_fecha(fecha)
        if not df_i.empty: i_lst.append(df_i)
    prog.empty()
    cat = lambda lst: pd.concat(lst, ignore_index=True) if lst else pd.DataFrame()
    return cat(m_lst), cat(c_lst), cat(i_lst)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 3 — GENERACION RDO
# ════════════════════════════════════════════════════════════════════════════════

def _tab3_fila(df):
    if df is None or df.empty: return []
    if df.shape[1] > 1:
        return df.iloc[:, 1:].apply(pd.to_numeric, errors="coerce").sum(axis=1, skipna=True).tolist()
    return []

def _tab3_totales_rer(df, barras):
    if df is None or df.empty: return []
    columnas = {str(c).strip().upper(): c for c in df.columns}
    sel = [columnas[str(b).strip().upper()] for b in barras if str(b).strip().upper() in columnas]
    if not sel: return []
    return df[sel].apply(pd.to_numeric, errors="coerce").sum(axis=1, skipna=True).tolist()

def _tab3_pad48(vals):
    if not vals: return []
    vals = list(vals)
    if len(vals) < 48: vals = [0.0] * (48 - len(vals)) + vals
    return vals[:48]

def _tab3_suma(a, b):
    if not a or not b: return []
    n = min(len(a), len(b))
    return [(a[i] if pd.notna(a[i]) else 0.0) + (b[i] if pd.notna(b[i]) else 0.0) for i in range(n)]

def _tab3_agregar(salida, fecha, serie, recurso, valores):
    valores = _tab3_pad48(valores)
    if not valores: return
    for i, v in enumerate(valores):
        salida.append({"Fecha": str(fecha), "Serie": serie, "Periodo": i + 1,
                       "Hora": HORAS_GENERACION[i], "Recurso": recurso, "MW": v})

def _csv_zip(fecha_dt, tipo, letra, nombre):
    contenido = obtener_zip(fecha_dt, tipo, letra)
    data      = _buscar_en_zip(contenido, nombre)
    return _leer_csv(data) if data else None

def _tab3_hidro(fecha_dt, fecha, salida):
    dh = _csv_zip(fecha_dt, "PDO", None, ARCHIVO_HIDRO)
    dr = _csv_zip(fecha_dt, "PDO", None, ARCHIVO_RER)
    vh = _tab3_fila(dh); vr = _tab3_totales_rer(dr, BARRAS_HIDRO)
    if vh and vr:
        vals = _tab3_suma(_tab3_pad48(vh), _tab3_pad48(vr))
        if vals: _tab3_agregar(salida, fecha, "PDO", "HIDRO", vals)
    for letra in RDO_LETRAS_GENERACION:
        dh = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_HIDRO)
        dr = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_RER)
        vh = _tab3_fila(dh); vr = _tab3_totales_rer(dr, BARRAS_HIDRO)
        if not vh or not vr: continue
        vals = _tab3_suma(_tab3_pad48(vh), _tab3_pad48(vr))
        if vals: _tab3_agregar(salida, fecha, f"RDO {letra}", "HIDRO", vals)

def _tab3_eolica(fecha_dt, fecha, salida):
    barras = [x.upper() for x in BARRAS_EOLICA]
    df = _csv_zip(fecha_dt, "PDO", None, ARCHIVO_RER)
    vals = _tab3_totales_rer(df, barras)
    if vals: _tab3_agregar(salida, fecha, "PDO", "EOLICA", _tab3_pad48(vals))
    for letra in RDO_LETRAS_GENERACION:
        df   = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_RER)
        vals = _tab3_totales_rer(df, barras)
        if vals: _tab3_agregar(salida, fecha, f"RDO {letra}", "EOLICA", _tab3_pad48(vals))

def _tab3_solar(fecha_dt, fecha, salida):
    barras = [x.upper() for x in BARRAS_SOLAR]
    df = _csv_zip(fecha_dt, "PDO", None, ARCHIVO_RER)
    vals = _tab3_totales_rer(df, barras)
    if vals: _tab3_agregar(salida, fecha, "PDO", "SOLAR", _tab3_pad48(vals))
    for letra in RDO_LETRAS_GENERACION:
        df   = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_RER)
        vals = _tab3_totales_rer(df, barras)
        if not vals: continue
        if not any(pd.notna(v) and float(v) != 0 for v in vals): continue
        _tab3_agregar(salida, fecha, f"RDO {letra}", "SOLAR", _tab3_pad48(vals))

def _tab3_termica(fecha_dt, fecha, salida):
    barras = [x.upper() for x in BARRAS_TERMICA_RER]
    dt = _csv_zip(fecha_dt, "PDO", None, ARCHIVO_TERMICA)
    dr = _csv_zip(fecha_dt, "PDO", None, ARCHIVO_RER)
    vt = _tab3_fila(dt); vr = _tab3_totales_rer(dr, barras)
    if vt and vr:
        vals = _tab3_suma(_tab3_pad48(vt), _tab3_pad48(vr))
        if vals: _tab3_agregar(salida, fecha, "PDO", "TERMICA", vals)
    for letra in RDO_LETRAS_GENERACION:
        dt = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_TERMICA)
        dr = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_RER)
        vt = _tab3_fila(dt); vr = _tab3_totales_rer(dr, barras)
        if not vt or not vr: continue
        vals = _tab3_suma(_tab3_pad48(vt), _tab3_pad48(vr))
        if vals: _tab3_agregar(salida, fecha, f"RDO {letra}", "TERMICA", vals)

def tab3_procesar_fecha(fecha_dt):
    fecha = fecha_dt.date()
    filas = []
    for nombre, fn in [("HIDRO", _tab3_hidro), ("EOLICA", _tab3_eolica),
                        ("SOLAR", _tab3_solar), ("TERMICA", _tab3_termica)]:
        try:
            fn(fecha_dt, fecha, filas)
        except Exception:
            pass
    return pd.DataFrame(filas, columns=["Fecha","Serie","Periodo","Hora","Recurso","MW"])

def tab3_procesar_rango(fecha_inicio, fecha_fin, log):
    fechas = pd.date_range(fecha_inicio, fecha_fin, freq="D")
    log.write(f"GENERACION RDO — {fecha_inicio.date()} → {fecha_fin.date()} ({len(fechas)} días)")
    prog   = st.progress(0)
    tablas = []
    for i, fdt in enumerate(fechas, 1):
        prog.progress(i / len(fechas))
        df_f = tab3_procesar_fecha(fdt)
        if not df_f.empty:
            tablas.append(df_f)
    prog.empty()
    return pd.concat(tablas, ignore_index=True) if tablas else pd.DataFrame()



# ════════════════════════════════════════════════════════════════════════════════
# TAB 4 — CMG PROGRAMADO (PDO + RDO)  →  tabla CMG_RDO
# ════════════════════════════════════════════════════════════════════════════════
# Sale del mismo ZIP de YUPANA que ya se descarga para DEMANDA y GENERACION
# (archivo "CMg - Barra ($ por MWh).csv"), así que no se baja nada extra.

CMG_SHEET   = "CMG_RDO"
ARCHIVO_CMG = "CMg - Barra ($ por MWh).csv"
BARRAS_CMG  = ["SANTA ROSA 220 A", "MOQUEGUA 220", "ZORRITOS 220"]


def _norm_barra(txt) -> str:
    import re as _re, unicodedata as _ud
    s = _ud.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode()
    return _re.sub(r"\s+", " ", s).strip().upper()


def _cmg_hora(p: int) -> str:
    m = p * 30
    return "00:00" if m >= 1440 else f"{m // 60:02d}:{m % 60:02d}"


def _cmg_largo(df, fecha_d: str, serie: str) -> list:
    """CSV ancho (una columna por barra) → filas. ETAPA 1..n = últimas n medias horas."""
    if df is None or df.empty:
        return []
    buscadas = {_norm_barra(b) for b in BARRAS_CMG}
    cols = [c for c in df.columns[1:] if _norm_barra(c) in buscadas]
    df = df.tail(48).reset_index(drop=True)
    n = len(df)
    filas = []
    for c in cols:
        barra = " ".join(str(c).split())
        vals = pd.to_numeric(df[c], errors="coerce").tolist()
        for k, v in enumerate(vals):
            p = 48 - n + 1 + k
            filas.append({"Fecha": fecha_d, "Serie": serie, "Periodo": p, "Hora": _cmg_hora(p),
                          "Barra": barra, "CMG": None if pd.isna(v) else round(float(v), 4)})
    return filas


def tab4_procesar_fecha(fecha_dt) -> pd.DataFrame:
    fecha_d = str(fecha_dt.date())
    filas = []
    filas += _cmg_largo(_csv_zip(fecha_dt, "PDO", None, ARCHIVO_CMG), fecha_d, "PDO")
    for letra in RDO_LETRAS_GENERACION:
        df = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_CMG)
        if df is None:
            break                       # no hay más RDO ese día
        filas += _cmg_largo(df, fecha_d, f"RDO {letra}")
    return pd.DataFrame(filas, columns=["Fecha", "Serie", "Periodo", "Hora", "Barra", "CMG"])


def tab4_procesar_rango(fecha_inicio, fecha_fin, log):
    fechas = pd.date_range(fecha_inicio, fecha_fin, freq="D")
    log.write(f"CMG RDO — {fecha_inicio.date()} → {fecha_fin.date()} ({len(fechas)} días)")
    prog = st.progress(0)
    tablas = []
    for i, fdt in enumerate(fechas, 1):
        prog.progress(i / len(fechas))
        df_f = tab4_procesar_fecha(fdt)
        if not df_f.empty:
            tablas.append(df_f)
    prog.empty()
    return pd.concat(tablas, ignore_index=True) if tablas else pd.DataFrame()


# ════════════════════════════════════════════════════════════════════════════════
# TABS 7 y 8 — TÉRMICAS GAS Y DIÉSEL POR CENTRAL  →  tabla TERMICAS_RDO
# ════════════════════════════════════════════════════════════════════════════════
# Sale del mismo ZIP de YUPANA ("Termica - Despacho (MW).csv"), no se baja nada extra.
# Una fila por Fecha + Serie + Periodo, con una columna por central (MW).
# Cada central suma sus unidades, igual que los tabs 7 y 8 de la app. Las unidades
# se reconocen por su nombre sin espacios (MAL1TG6GAS y "MALACAS1 TG 6  GAS" caen
# en MALACAS gas); las que no son de ninguna central van a "otras".
# Las columnas deben coincidir con Supabase y con la tabla TERMICAS_CENTRALES.

TERMICAS_SHEET = "TERMICAS_RDO"

_PREFIJOS_CENTRAL = [
    ("KALLPA", "kallpa"), ("CHILCA1", "chilca1"), ("CHILCA2", "chilca2"),
    ("FENIX", "fenix"), ("VENT", "ventanilla"), ("OLLEROS", "olleros"),
    ("LFLORES", "las_flores"), ("LASFLORES", "las_flores"),
    ("INDEP", "independencia"), ("STAROSA", "santa_rosa"), ("SANTAROSA", "santa_rosa"),
    ("MALACAS", "malacas"), ("MAL", "malacas"),
    ("AGUAYTIA", "aguaytia"), ("AGE", "aguaytia"),
    ("RFILO", "rf_ilo"), ("RECKA", "recka"),
    ("PTOBRVO", "puerto_bravo"), ("PUERTOBRAVO", "puerto_bravo"),
    ("CTNEPI", "ilo4"), ("NEPI", "ilo4"), ("RFETEN", "eten"),
    ("RFPTOMALDONADO", "rf_pto_maldonado"), ("RFPUCALLPA", "rf_pucallpa"),
    ("CHILINA", "chilina"), ("TUMBES", "tumbes"),
]
TERMICAS_COLS_MW = [
    "g_kallpa", "g_chilca1", "g_chilca2", "g_fenix", "g_ventanilla", "g_olleros",
    "g_las_flores", "g_independencia", "g_santa_rosa", "g_malacas", "g_aguaytia", "g_otras",
    "d_malacas", "d_ventanilla", "d_rf_ilo", "d_recka", "d_puerto_bravo", "d_fenix",
    "d_ilo4", "d_eten", "d_rf_pto_maldonado", "d_rf_pucallpa", "d_chilina",
    "d_santa_rosa", "d_tumbes", "d_otras",
    "o_otras",
]
TERMICAS_COLUMNAS = ["Fecha", "Serie", "Periodo"] + TERMICAS_COLS_MW


def _termica_columna(nombre) -> str:
    """'KALLPATG1GAS' -> 'g_kallpa', 'FENIX GT11  D2' -> 'd_fenix', 'SNTV1R500' -> 'o_otras'."""
    import re as _re, unicodedata as _ud
    s = _ud.normalize("NFKD", str(nombre)).encode("ASCII", "ignore").decode().upper()
    s = _re.sub(r"\s+", "", s)
    if "GAS" in s:
        pre = "g_"
    elif "D2" in s or "DIESEL" in s:
        pre = "d_"
    else:
        return "o_otras"
    clave = next((c for p, c in _PREFIJOS_CENTRAL if s.startswith(p)), "otras")
    col = pre + clave
    return col if col in TERMICAS_COLS_MW else pre + "otras"


def _termicas_por_central(df, fecha_d: str, serie: str) -> pd.DataFrame:
    """CSV ancho (una columna por unidad) -> una fila por periodo, una columna por central."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.tail(48).reset_index(drop=True)
    n = len(df)
    out = pd.DataFrame(0.0, index=range(n), columns=TERMICAS_COLS_MW)
    for c in df.columns[1:]:
        out[_termica_columna(c)] += pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    out = out.round(3)
    out.insert(0, "Periodo", range(48 - n + 1, 49))
    out.insert(0, "Serie", serie)
    out.insert(0, "Fecha", fecha_d)
    return out[TERMICAS_COLUMNAS]


def tab78_procesar_fecha(fecha_dt) -> pd.DataFrame:
    fecha_d = str(fecha_dt.date())
    partes = []
    df = _csv_zip(fecha_dt, "PDO", None, ARCHIVO_TERMICA)
    if df is not None and not df.empty:
        partes.append(_termicas_por_central(df, fecha_d, "PDO"))
    for letra in RDO_LETRAS_GENERACION:
        df = _csv_zip(fecha_dt, "RDO", letra, ARCHIVO_TERMICA)
        if df is None:
            break                       # no hay más RDO ese día
        if not df.empty:
            partes.append(_termicas_por_central(df, fecha_d, f"RDO {letra}"))
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=TERMICAS_COLUMNAS)


def tab78_procesar_rango(fecha_inicio, fecha_fin, log):
    fechas = pd.date_range(fecha_inicio, fecha_fin, freq="D")
    log.write(f"TÉRMICAS por central — {fecha_inicio.date()} → {fecha_fin.date()} ({len(fechas)} días)")
    prog = st.progress(0)
    tablas = []
    for i, fdt in enumerate(fechas, 1):
        prog.progress(i / len(fechas))
        try:
            df_f = tab78_procesar_fecha(fdt)
        except Exception as e:
            log.write(f"  ⚠️ {fdt.date()} térmicas: {e}")
            continue
        if not df_f.empty:
            tablas.append(df_f)
    prog.empty()
    return pd.concat(tablas, ignore_index=True) if tablas else pd.DataFrame()


# ════════════════════════════════════════════════════════════════════════════════
# TABS 5 y 6 — IEOD EJECUTADO  →  tablas IEOD_TIPO_RECURSO e IEOD_G_AREAS
# ════════════════════════════════════════════════════════════════════════════════
# Archivo AnexoA_ddmm.xlsx del IEOD (sale al día siguiente). Se leen dos hojas
# en la misma descarga:
#   TIPO_RECURSO → generación por tipo de recurso / combustible  (tab 5 y 6)
#   G_AREAS      → generación por área Norte/Centro/Sur          (tab 6: eólica N/C)
# La vista HISTORICO_RECURSO de Supabase (IEOD o, si no hay, PDO+RDO) se
# actualiza sola a partir de estas tablas y de GENERACION_RDO.
#
# Las funciones de lectura son las mismas del script ieod_tipo_recurso.py:
# reconocen columnas por palabra clave (con/sin tildes, "+", "&"), aguantan
# encabezados vacíos, columnas repetidas y columnas nuevas.

IEOD_TIPO_SHEET  = "IEOD_TIPO_RECURSO"
IEOD_AREAS_SHEET = "IEOD_G_AREAS"
IEOD_TOLERANCIA_TOTAL_MW = 1.0

_BASE_IEOD = (
    "https://www.coes.org.pe/portal/browser/download?"
    "url=Post%20Operaci%C3%B3n%2FReportes%2FIEOD%2F"
    "{y}%2F{m}_{M}%2F{d}%2FAnexoA_{d}{m}.xlsx"
)

def _url_ieod(fecha):
    y, m, d = fecha.year, f"{fecha.month:02d}", f"{fecha.day:02d}"
    return _BASE_IEOD.format(y=y, m=m, M=_mes(fecha.month), d=d)

from datetime import date, datetime, time as dtime   # usados por las funciones IEOD
import re
import unicodedata

IEOD_HOJA_TIPO = "TIPO_RECURSO"

IEOD_HOJA_AREAS = "G_AREAS"

REGLAS = [
    (["PASADA"],                "H. PASADA",                   "HIDRO"),
    (["REGULACION"],            "H. REGULACION",               "HIDRO"),
    (["CAMISEA"],               "GAS DE CAMISEA",              "TERMICA"),
    (["MALACAS"],               "GAS DE MALACAS",              "TERMICA"),
    (["AGUAYTIA"],              "GAS DE AGUAYTIA",             "TERMICA"),
    (["NAFTA", "FLEXIGAS"],     "NAFTA Y FLEXIGAS",            "TERMICA"),   # columna combinada
    (["NAFTA"],                 "NAFTA Y GAS DE REFINERIA",    "TERMICA"),
    (["FLEXIGAS"],              "FLEXIGAS Y GAS DE REFINERIA", "TERMICA"),
    (["RESIDUAL"],              "RESIDUAL",                    "TERMICA"),
    (["DIESEL"],                "DIESEL",                      "TERMICA"),
    (["BAGA"],                  "BAGAZO",                      "TERMICA"),   # BAGAZO / BAGASO
    (["BIOGAS"],                "BIOGAS",                      "TERMICA"),
    (["CARBON"],                "CARBON",                      "TERMICA"),
    (["SOLAR"],                 "SOLAR",                       "SOLAR"),
    (["EOLICA"],                "EOLICA",                      "EOLICA"),
    (["TOTAL"],                 "TOTAL",                       "DEMANDA"),
]

PALABRAS_TERMICAS = ["GAS", "PETROLEO", "CARBON", "DIESEL", "RESIDUAL", "BIOMASA", "COMBUST"]

COLUMNAS_TIPO = ["Fecha", "Periodo", "Hora", "Columna", "Grupo", "MW", "Columna_IEOD"]

COLUMNAS_AREAS = ["Fecha", "Periodo", "Hora", "Bloque", "Area", "Tipo", "Columna", "MW", "Columna_IEOD"]

columnas_nuevas_vistas: dict[str, str] = {}

def limpiar(txt) -> str:
    """Quita tildes, saltos de línea y espacios dobles; pasa a mayúsculas."""
    s = unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode()
    s = s.replace("&", " Y ").replace("+", " Y ")
    return re.sub(r"\s+", " ", s).strip().upper()

def reconocer(encabezado: str) -> tuple[str, str]:
    """Devuelve (nombre unificado, grupo) para un encabezado del IEOD."""
    t = limpiar(encabezado)
    palabras = set(re.findall(r"[A-Z0-9]+", t))
    for claves, nombre, grupo in REGLAS:
        if all(any(p.startswith(c) for p in palabras) for c in claves):
            return nombre, grupo
    grupo = "TERMICA" if any(k in t for k in PALABRAS_TERMICAS) else "OTRO"
    return t, grupo

def a_hora(valor) -> str | None:
    """Convierte la celda de hora ('00:30', time, datetime, 0.0208…) a 'HH:MM'."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (datetime, pd.Timestamp)):
        valor = valor.time()
    if isinstance(valor, dtime):
        return f"{valor.hour:02d}:{valor.minute:02d}"
    if isinstance(valor, (int, float)) and 0 <= valor <= 1:          # fracción de día de Excel
        minutos = round(valor * 1440)
        return f"{(minutos // 60) % 24:02d}:{minutos % 60:02d}"
    m = re.match(r"^\s*(\d{1,2}):(\d{2})", str(valor))
    return f"{int(m.group(1)):02d}:{m.group(2)}" if m else None

def periodo_de_hora(h: str) -> int:
    hh, mm = map(int, h.split(":"))
    minutos = hh * 60 + mm
    if minutos == 0 or (hh == 23 and mm == 59):
        return 48
    return round(minutos / 30)

def hora_de_periodo(p: int) -> str:
    minutos = p * 30
    return "00:00" if minutos >= 1440 else f"{minutos // 60:02d}:{minutos % 60:02d}"

def leer_tipo_recurso(xls: pd.ExcelFile, f: date) -> tuple[pd.DataFrame | None, list[str]]:
    """Hoja TIPO_RECURSO. Devuelve (filas en formato largo, avisos)."""
    avisos: list[str] = []
    hoja = next((h for h in xls.sheet_names if limpiar(h) == IEOD_HOJA_TIPO), None)
    if hoja is None:
        return None, [f"no tiene la hoja {IEOD_HOJA_TIPO}"]
    raw = pd.read_excel(xls, sheet_name=hoja, header=None)

    # 1) Fila de encabezados: la primera que tenga una celda "HORA"
    fila_hdr = col_hora = None
    for i in range(min(40, len(raw))):
        for j, v in enumerate(raw.iloc[i].tolist()):
            if limpiar(v) == "HORA":
                fila_hdr, col_hora = i, j
                break
        if fila_hdr is not None:
            break
    if fila_hdr is None:
        return None, ["no se encontró la fila de encabezados (HORA)"]

    # 2) Filas de datos: debajo del encabezado, las que tienen una hora válida
    filas = []
    for i in range(fila_hdr + 1, min(fila_hdr + 80, len(raw))):
        h = a_hora(raw.iat[i, col_hora])
        if h is None:
            if filas:                       # se terminó el bloque (TOTAL MWh, %, …)
                break
            continue
        filas.append((i, periodo_de_hora(h)))
    if len(filas) != 48:
        avisos.append(f"tiene {len(filas)} medias horas (se esperaban 48)")

    def vacio(v) -> bool:
        return v is None or (isinstance(v, float) and pd.isna(v)) or not str(v).strip()

    def tiene_datos(j: int) -> bool:
        return any(not pd.isna(pd.to_numeric(raw.iat[i, j], errors="coerce")) for i, _ in filas)

    # 3) Columnas de datos: desde HORA hasta la columna TOTAL (incluida).
    #    A la derecha hay otra tabla (EMPRESA/CENTRAL) que no se toma.
    #    El COES a veces deja encabezados vacíos en medio, por eso NO se corta
    #    en el primer vacío: se recorre hasta TOTAL.
    j_total = None
    for j in range(col_hora + 1, raw.shape[1]):
        t = limpiar(raw.iat[fila_hdr, j])
        if t in ("EMPRESA", "CENTRAL"):
            break
        if t == "TOTAL" or t.startswith("TOTAL "):
            j_total = j
            break
    if j_total is None:
        # sin TOTAL: se toma hasta EMPRESA o hasta 3 columnas seguidas totalmente vacías
        avisos.append("no tiene columna TOTAL")
        j_fin, vacias = col_hora, 0
        for j in range(col_hora + 1, raw.shape[1]):
            t = limpiar(raw.iat[fila_hdr, j])
            if t in ("EMPRESA", "CENTRAL"):
                break
            if vacio(raw.iat[fila_hdr, j]) and not tiene_datos(j):
                vacias += 1
                if vacias >= 3:
                    break
                continue
            vacias, j_fin = 0, j
        j_total_rango = j_fin
    else:
        j_total_rango = j_total

    cols: list[tuple[int, str]] = []
    for j in range(col_hora + 1, j_total_rango + 1):
        v = raw.iat[fila_hdr, j]
        if vacio(v):
            if tiene_datos(j):
                nombre_col = f"SIN ENCABEZADO (col {j + 1})"
                cols.append((j, nombre_col))
                avisos.append(f"la columna {j + 1} tiene datos pero no tiene nombre → '{nombre_col}' (grupo OTRO)")
            continue                         # encabezado vacío y sin datos: se ignora
        cols.append((j, re.sub(r"\s+", " ", str(v)).strip()))
    if not cols:
        return None, ["no se encontraron columnas de datos"]

    registros = []
    usados: dict[str, int] = {}
    for j, encabezado in cols:
        if encabezado.startswith("SIN ENCABEZADO"):
            nombre, grupo = encabezado, "OTRO"
        else:
            nombre, grupo = reconocer(encabezado)
            if grupo == "OTRO" or nombre not in {r[1] for r in REGLAS}:
                if nombre not in columnas_nuevas_vistas:
                    columnas_nuevas_vistas[nombre] = f"{f:%d/%m/%Y}"
                    avisos.append(f"columna NUEVA '{encabezado}' → guardada como '{nombre}' (grupo {grupo})")
        # Columna repetida el mismo día (p. ej. "DIESEL" dos veces): se guardan ambas
        usados[nombre] = usados.get(nombre, 0) + 1
        if usados[nombre] > 1:
            nuevo = f"{nombre} ({usados[nombre]})"
            avisos.append(f"columna repetida '{encabezado}' → guardada como '{nuevo}'")
            nombre = nuevo
        for i, p in filas:
            mw = pd.to_numeric(raw.iat[i, j], errors="coerce")
            registros.append((f.isoformat(), p, hora_de_periodo(p), nombre, grupo,
                              None if pd.isna(mw) else round(float(mw), 4), encabezado))

    df = pd.DataFrame(registros, columns=COLUMNAS_TIPO)

    # 4) Control: suma de columnas vs TOTAL
    if (df["Columna"] == "TOTAL").any():
        piv = df.pivot_table(index="Periodo", columns="Columna", values="MW", aggfunc="sum").fillna(0)
        dif = (piv.drop(columns="TOTAL").sum(axis=1) - piv["TOTAL"]).abs()
        if dif.max() > IEOD_TOLERANCIA_TOTAL_MW:
            p = int(dif.idxmax())
            avisos.append(f"la suma de columnas no cuadra con TOTAL (hasta {dif.max():.1f} MW a las {hora_de_periodo(p)})")
    else:
        avisos.append("no tiene columna TOTAL")

    return df, avisos

def leer_g_areas(xls: pd.ExcelFile, f: date) -> tuple[pd.DataFrame | None, list[str]]:
    """Hoja G_AREAS (generación por área y tipo). Devuelve (filas en formato largo, avisos)."""
    avisos: list[str] = []
    hoja = next((h for h in xls.sheet_names if limpiar(h) == IEOD_HOJA_AREAS), None)
    if hoja is None:
        return None, [f"no tiene la hoja {IEOD_HOJA_AREAS}"]
    raw = pd.read_excel(xls, sheet_name=hoja, header=None)
    n_filas, n_cols = raw.shape

    def vacio(v) -> bool:
        return v is None or (isinstance(v, float) and pd.isna(v)) or not str(v).strip()

    # 1) Fila de encabezados: la que más celdas tiene con HIDRO/TERMO/SOLAR/EOLICA/TOTAL
    claves = ("HIDRO", "TERMO", "SOLAR", "EOLICA", "TOTAL")
    mejor, fila_hdr = 0, None
    for i in range(min(30, n_filas)):
        n = sum(any(k in limpiar(v) for k in claves) for v in raw.iloc[i].tolist() if not vacio(v))
        if n > mejor:
            mejor, fila_hdr = n, i
    if fila_hdr is None or mejor < 3:
        return None, ["no se encontró la fila de encabezados"]

    # 2) Columna de HORA (la primera con horas válidas debajo del encabezado) y filas de datos
    col_hora, filas = None, []
    for j in range(n_cols):
        cand = []
        for i in range(fila_hdr + 1, min(fila_hdr + 80, n_filas)):
            h = a_hora(raw.iat[i, j])
            if h is None:
                if cand:
                    break
                continue
            cand.append((i, periodo_de_hora(h)))
        if len(cand) >= 24:
            col_hora, filas = j, cand
            break
    if col_hora is None:
        return None, ["no se encontró la columna HORA"]
    if len(filas) != 48:
        avisos.append(f"tiene {len(filas)} medias horas (se esperaban 48)")

    # 3) Rótulos de bloque/área de las filas de arriba (rellenados hacia la derecha
    #    por si vienen como celdas combinadas)
    def rotulos(i: int) -> list[str]:
        out, ultimo = [], ""
        for j in range(n_cols):
            v = raw.iat[i, j] if 0 <= i < n_filas else None
            if not vacio(v):
                ultimo = limpiar(v)
            if not vacio(v) and limpiar(v) == "HORA":
                ultimo = ""
            out.append(ultimo)
        return out
    rot1 = rotulos(fila_hdr - 1)     # NORTE / CENTRO / SUR / INTEGRANTES COES / ... RER
    rot2 = rotulos(fila_hdr - 2)     # RESUMEN ... POR AREAS / GENERACION POR TIPO ...

    registros, usados = [], {}
    for j in range(col_hora + 1, n_cols):
        enc = raw.iat[fila_hdr, j]
        if vacio(enc):
            continue                                    # columnas HORA intermedias, separadores
        t = limpiar(enc)
        if t == "HORA":
            continue
        contexto = f"{t} {rot1[j]} {rot2[j]}"
        if "RER" in t or "RER" in rot1[j]:
            bloque = "RER"
        elif "AREA" in rot2[j]:
            bloque = "AREAS"
        else:
            bloque = "TIPO"
        palabras = t.split()      # por palabra completa: "GEOTERMICA" no debe pasar por "TERMICA"
        if any(w.startswith("HIDRO") for w in palabras):
            tipo = "HIDRO"
        elif any(w.startswith("TERMO") or w in ("TERMICA", "TERMICAS") for w in palabras):
            tipo = "TERMICA"
        elif any(w.startswith("SOLAR") for w in palabras):
            tipo = "SOLAR"
        elif any(w.startswith("EOLIC") for w in palabras):
            tipo = "EOLICA"
        elif "TOTAL" in palabras or "GENERACION COES" in t:
            tipo = "TOTAL"
        else:
            tipo = "OTRO"
        area = next((a for a in ("NORTE", "CENTRO", "SUR") if a in t.split()), "SEIN")
        nombre = f"{tipo} {area}"
        if tipo == "OTRO":
            nombre = t
            clave_nueva = f"G_AREAS: {t}"
            if clave_nueva not in columnas_nuevas_vistas:
                columnas_nuevas_vistas[clave_nueva] = f"{f:%d/%m/%Y}"
                avisos.append(f"columna NUEVA '{enc}' → guardada como '{t}' (bloque {bloque}, tipo OTRO)")
        k = (bloque, nombre)
        usados[k] = usados.get(k, 0) + 1
        if usados[k] > 1:
            avisos.append(f"columna repetida '{enc}' en bloque {bloque} → '{nombre} ({usados[k]})'")
            nombre = f"{nombre} ({usados[k]})"
        enc_txt = re.sub(r"\s+", " ", str(enc)).strip()
        for i, p in filas:
            mw = pd.to_numeric(raw.iat[i, j], errors="coerce")
            registros.append((f.isoformat(), p, hora_de_periodo(p), bloque, area, tipo, nombre,
                              None if pd.isna(mw) else round(float(mw), 4), enc_txt))

    if not registros:
        return None, avisos + ["no se encontraron columnas de datos"]
    df = pd.DataFrame(registros, columns=COLUMNAS_AREAS)

    # 4) Controles: en cada bloque, la suma de tipos debe dar el TOTAL SEIN
    for bloque, sub in df.groupby("Bloque"):
        tot = sub[sub["Columna"] == "TOTAL SEIN"]
        if tot.empty:
            continue
        partes = sub[(sub["Tipo"] != "TOTAL") & (sub["Columna"] != "TOTAL SEIN")]
        s1 = partes.groupby("Periodo")["MW"].sum()
        s2 = tot.groupby("Periodo")["MW"].sum()
        dif = (s1.reindex(s2.index, fill_value=0) - s2).abs()
        if len(dif) and dif.max() > IEOD_TOLERANCIA_TOTAL_MW:
            p = int(dif.idxmax())
            avisos.append(f"bloque {bloque}: la suma no cuadra con TOTAL "
                          f"(hasta {dif.max():.1f} MW a las {hora_de_periodo(p)})")
    return df, avisos


def tab56_procesar_fecha(fecha_dt):
    """Devuelve (df_tipo, df_areas, avisos). None si el IEOD de ese día aún no existe."""
    data = _get(_url_ieod(fecha_dt), timeout=90)
    if not data or data[:4] != b"PK\x03\x04":
        return None, None, ["IEOD aún no publicado"]
    f = fecha_dt.date()
    avisos = []
    try:
        xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    except Exception as e:
        return None, None, [f"no se pudo abrir el IEOD: {e}"]
    try:
        df_t, av1 = leer_tipo_recurso(xls, f)
        avisos += [f"TIPO_RECURSO: {a}" for a in av1]
    except Exception as e:
        df_t = None; avisos.append(f"TIPO_RECURSO: error {e}")
    try:
        df_a, av2 = leer_g_areas(xls, f)
        avisos += [f"G_AREAS: {a}" for a in av2]
    except Exception as e:
        df_a = None; avisos.append(f"G_AREAS: error {e}")
    return df_t, df_a, avisos


IEOD_AREAS_COLUMNAS_SHEET = "IEOD_G_AREAS_COLUMNAS"

def _actualizar_columnas_g_areas(df_a, con):
    """
    Area, Tipo y Columna_IEOD dependen solo de (Bloque, Columna), así que se guardan
    una sola vez en IEOD_G_AREAS_COLUMNAS (Power BI las vuelve a unir al cargar).
    Si el COES agrega una columna nueva, aquí se registra sola.
    """
    if df_a is None or df_a.empty or not _columnas_tabla(con, IEOD_AREAS_COLUMNAS_SHEET):
        return
    cols = ["Bloque", "Columna", "Area", "Tipo", "Columna_IEOD"]
    if any(c not in df_a.columns for c in cols):
        return
    dim = df_a[cols].drop_duplicates(subset=["Bloque", "Columna"])
    dim = dim.astype(object).where(pd.notnull(dim), None)
    cur = con.cursor()
    try:
        psycopg2.extras.execute_values(
            cur,
            f'INSERT INTO "{IEOD_AREAS_COLUMNAS_SHEET}" ("Bloque","Columna","Area","Tipo","Columna_IEOD") '
            'VALUES %s ON CONFLICT ("Bloque","Columna") DO NOTHING',
            [tuple(r) for r in dim.itertuples(index=False)])
        con.commit()
    except Exception:
        con.rollback()
        raise


def tab56_procesar_rango(fecha_inicio, fecha_fin, log):
    fechas = pd.date_range(fecha_inicio, fecha_fin, freq="D")
    log.write(f"IEOD (TIPO_RECURSO + G_AREAS) — {fecha_inicio.date()} → {fecha_fin.date()} ({len(fechas)} días)")
    prog = st.progress(0)
    lt, la = [], []
    for i, fdt in enumerate(fechas, 1):
        prog.progress(i / len(fechas))
        df_t, df_a, avisos = tab56_procesar_fecha(fdt)
        if df_t is not None and not df_t.empty:
            lt.append(df_t)
        if df_a is not None and not df_a.empty:
            la.append(df_a)
        for a in avisos:
            if a == "IEOD aún no publicado":
                log.write(f"  {fdt.date()}: IEOD aún no publicado (se intentará en la próxima actualización)")
            else:
                log.write(f"  ⚠️ {fdt.date()} {a}")
    prog.empty()
    cat = lambda lst: pd.concat(lst, ignore_index=True) if lst else pd.DataFrame()
    return cat(lt), cat(la)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 10 — POTENCIA ACTIVA (medidores, Excel mensual)  →  tabla POTENCIA_ACTIVA
# ════════════════════════════════════════════════════════════════════════════════
# No se descarga del COES: es el Excel mensual "REGISTROS DE MEDIDORES EN BORNES DE
# GENERACIÓN CADA 15 MINUTOS DE POTENCIA ACTIVA (MW)" que alguien sube a la app.
# Se guardan las centrales solares (C.S.) y eólicas (C.E.), con sus unidades/bloques
# sumados por fecha (igual que el tab 10). p01 = 00:15 ... p96 = 00:00.

POTENCIA_SHEET = "POTENCIA_ACTIVA"
POTENCIA_CENTRALES_SHEET = "POTENCIA_ACTIVA_CENTRALES"
POTENCIA_COLS_P = [f"p{i:02d}" for i in range(1, 97)]


def _pa_sin_tildes(s) -> str:
    import unicodedata as _ud
    return _ud.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode()


def _pa_norm_central(s) -> str:
    import re as _re
    return _re.sub(r"\s+", " ", str(s)).strip().upper()


def _pa_hora_a_indice(h):
    import re as _re
    m = _re.match(r"^\s*(\d{1,2}):(\d{2})\s*$", str(h))
    if not m:
        return None
    minutos = int(m.group(1)) * 60 + int(m.group(2))
    if minutos == 0:
        return 96
    if minutos % 15:
        return None
    return minutos // 15


def leer_excel_potencia(data: bytes, nombre: str = "") -> pd.DataFrame:
    """Excel de medidores -> Fecha | Central | p01..p96 (solo C.S. y C.E., unidades sumadas)."""
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    hoja = next((h for h in xls.sheet_names if "ACTIVA" in _pa_sin_tildes(h).upper()), xls.sheet_names[0])
    crudo = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=40)
    fila = 9
    for i in range(len(crudo)):
        vals = {_pa_sin_tildes(v).strip().upper() for v in crudo.iloc[i].dropna().astype(str)}
        if "FECHA" in vals and "CENTRAL" in vals:
            fila = i
            break
    df = pd.read_excel(xls, sheet_name=hoja, header=fila)
    df.columns = [str(c).strip() for c in df.columns]
    col = {_pa_sin_tildes(c).upper(): c for c in df.columns}
    c_fecha, c_central = col.get("FECHA"), col.get("CENTRAL")
    if not c_fecha or not c_central:
        raise ValueError(f"{nombre}: no encontré las columnas FECHA y CENTRAL")
    horas = {}
    for c in df.columns:
        k = _pa_hora_a_indice(c)
        if k is not None and k not in horas.values():
            horas[c] = k
    if len(horas) < 90:
        raise ValueError(f"{nombre}: solo encontré {len(horas)} columnas de hora (se esperaban 96)")
    fechas = pd.to_datetime(df[c_fecha].astype(str).str.strip(), format="%d/%m/%Y", errors="coerce")
    df = df.assign(_fecha=fechas, _central=df[c_central].map(_pa_norm_central))
    df = df[df["_fecha"].notna()]                                   # fuera totales y leyenda
    df = df[df["_central"].str.match(r"^C\.(S|E)\.")]             # solo solares y eólicas
    valores = df[list(horas)].apply(pd.to_numeric, errors="coerce")
    valores.columns = [f"p{horas[c]:02d}" for c in horas]
    valores["_fecha"] = df["_fecha"].dt.strftime("%Y-%m-%d")
    valores["_central"] = df["_central"]
    suma = valores.groupby(["_fecha", "_central"]).sum(min_count=1).reset_index()
    suma = suma.rename(columns={"_fecha": "Fecha", "_central": "Central"})
    for c in POTENCIA_COLS_P:
        if c not in suma.columns:
            suma[c] = None
    suma[POTENCIA_COLS_P] = suma[POTENCIA_COLS_P].round(5)
    return suma[["Fecha", "Central"] + POTENCIA_COLS_P]


def subir_potencia_pg(df: pd.DataFrame, con) -> int:
    """Agrega centrales nuevas a la lista y reemplaza por fecha las filas de POTENCIA_ACTIVA."""
    if df is None or df.empty:
        return 0
    cur = con.cursor()
    try:
        centrales = sorted(df["Central"].unique())
        psycopg2.extras.execute_values(
            cur,
            f'INSERT INTO "{POTENCIA_CENTRALES_SHEET}" ("Central","Grupo","Tipo","Orden","Tab10") VALUES %s '
            'ON CONFLICT ("Central") DO NOTHING',
            [(c, c, "SOLAR" if c.startswith("C.S.") else "EOLICA", 99, False) for c in centrales])
        con.commit()
    except Exception:
        con.rollback()
        raise
    return _insertar_df_pg(df, POTENCIA_SHEET, con)


# ════════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Actualización de Datos COES", page_icon=None, layout="centered")

# ── Estilos modernos ──────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Fondo general — azul oscuro Osinergmin */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(160deg, #001a5e, #0039AA, #001233);
    min-height: 100vh;
    overflow: hidden;
}
[data-testid="stHeader"] { background: transparent; }

/* Esferas animadas — movimiento vertical */
[data-testid="stAppViewContainer"]::before,
[data-testid="stAppViewContainer"]::after {
    content: "";
    position: fixed;
    border-radius: 50%;
    filter: blur(65px);
    opacity: 0.45;
    pointer-events: none;
    z-index: 0;
}
[data-testid="stAppViewContainer"]::before {
    width: 650px; height: 650px;
    background: radial-gradient(circle, #FCE122, #E08A00);
    top: -200px; left: -150px;
    animation: floatUp 9s ease-in-out infinite alternate;
}
[data-testid="stAppViewContainer"]::after {
    width: 550px; height: 550px;
    background: radial-gradient(circle, #0039AA, #12996B);
    bottom: -150px; right: -120px;
    animation: floatDown 11s ease-in-out infinite alternate;
}

@keyframes floatUp {
    0%   { transform: translateY(0px) scale(1); }
    50%  { transform: translateY(80px) scale(1.06); }
    100% { transform: translateY(160px) scale(0.96); }
}
@keyframes floatDown {
    0%   { transform: translateY(0px) scale(1); }
    50%  { transform: translateY(-80px) scale(1.05); }
    100% { transform: translateY(-160px) scale(0.97); }
}

/* Puntitos decorativos */
[data-testid="stMain"] {
    background-image: radial-gradient(rgba(255,255,255,0.05) 1px, transparent 1px);
    background-size: 28px 28px;
    position: relative;
    z-index: 1;
}

/* Título */
.titulo-app {
    text-align: center;
    font-size: 2.4rem;
    font-weight: 800;
    color: #ffffff;
    margin-bottom: 0.2rem;
    letter-spacing: -0.5px;
}
.subtitulo-app {
    text-align: center;
    color: #c3cbe6;
    font-size: 0.95rem;
    margin-bottom: 2rem;
}

/* Tarjetas de estado */
.card-estado {
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 14px;
    padding: 1rem 1.3rem;
    margin-bottom: 0.6rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
    backdrop-filter: blur(8px);
}
.card-tabla  { font-weight: 700; color: #ffffff; font-size: 0.95rem; }
.card-fecha  { color: #c3cbe6; font-size: 0.85rem; }
.badge-ok, .badge-warn, .badge-err, .badge-info {
    border-radius: 20px; padding: 4px 14px; font-size: 0.8rem; font-weight: 700;
    white-space: nowrap; margin-left: 12px;
}
.badge-ok    { background: rgba(18,153,107,0.30); color: #6CF5BC; border: 1px solid #6CF5BC88; }
.badge-warn  { background: rgba(252,225,34,0.18); color: #FFE95C; border: 1px solid #FFE95C88; }
.badge-err   { background: rgba(209,67,67,0.30);  color: #FF9B9B; border: 1px solid #FF9B9B88; }
.badge-info  { background: rgba(140,175,255,0.22); color: #D6E2FF; border: 1px solid #D6E2FF66; }

/* Tab 10: meses cargados como fichas */
.pa-meses { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 0.6rem; }
.pa-mes   { border-radius: 8px; padding: 2px 9px; font-size: 0.75rem; font-weight: 700; }
.pa-ok    { background: rgba(18,153,107,0.35); color: #E8FFF5; border: 1px solid #6CF5BC99; }
.pa-inc   { background: rgba(252,225,34,0.25); color: #FFF6B0; border: 1px solid #FFE95C99; }
.pa-falta { background: rgba(255,255,255,0.06); color: #c3cbe6; border: 1px dashed #c3cbe699; }
.pa-leyenda { color: #c3cbe6; font-size: 0.78rem; margin-top: 0.5rem; }
.card-col { display: block; }
.card-col .card-cab { display: flex; justify-content: space-between; align-items: center; }

/* Mensajes propios (se leen sobre el fondo azul) */
.msg-linea {
    background: rgba(255,255,255,0.10); border-left: 3px solid #6CF5BC;
    border-radius: 6px; padding: 0.5rem 1rem; margin-bottom: 0.4rem;
    color: #ffffff; font-size: 0.88rem;
}
.msg-linea.err { border-left-color: #FF9B9B; }

/* Panel de subida del tab 10 (st.container con borde que contiene .panel-titulo) */
/* Solo el panel (no el bloque principal de la página, que también tiene envoltorio) */
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] > [data-testid="stElementContainer"] .panel-titulo),
[data-testid="stVerticalBlock"]:not([data-testid="stVerticalBlockBorderWrapper"] > div > [data-testid="stVerticalBlock"]):has(> [data-testid="stElementContainer"] .panel-titulo) {
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.18) !important;
    border-radius: 14px;
    backdrop-filter: blur(8px);
}
/* Botón deshabilitado (sin archivos): que se lea sobre el fondo */
button[kind="primary"]:disabled, button[data-testid="stBaseButton-primary"]:disabled {
    background: rgba(255,255,255,0.12) !important;
    color: #c3cbe6 !important;
    border: 1px solid rgba(255,255,255,0.30) !important;
    opacity: 1 !important;
}
button[kind="primary"]:disabled *, button[data-testid="stBaseButton-primary"]:disabled * { color: #c3cbe6 !important; }
.panel-titulo { color: #ffffff; font-weight: 700; font-size: 0.95rem; margin-bottom: 0.2rem; }
.panel-ayuda  { color: #c3cbe6; font-size: 0.82rem; margin-bottom: 0.4rem; }

/* File uploader: zona blanca y archivos elegidos legibles */
[data-testid="stFileUploaderDropzone"] { background: #ffffff; border-radius: 10px; }
[data-testid="stFileUploaderFile"] {
    background: #ffffff; border-radius: 8px; padding: 4px 10px; margin-top: 6px;
}
[data-testid="stFileUploaderFile"] * { color: #1b2140 !important; }
[data-testid="stFileUploaderPagination"] * { color: #e0e6ff !important; }

/* Avisos de Streamlit (st.error / st.success / st.warning) legibles sobre el azul */
[data-testid="stAlert"] * { color: #ffffff !important; }
[data-testid="stAlert"] > div { background: rgba(255,255,255,0.10); border: 1px solid rgba(255,255,255,0.25); }
[data-testid="stAlertContentError"] { border-left: 4px solid #FF9B9B; padding-left: 0.6rem; }
[data-testid="stAlertContentSuccess"] { border-left: 4px solid #6CF5BC; padding-left: 0.6rem; }

/* Spinner y textos sueltos de Streamlit sobre el fondo oscuro */
[data-testid="stSpinner"] * { color: #e0e6ff !important; }

/* Label del date_input */
label { color: #c3cbe6 !important; font-size: 0.85rem !important; }

/* Sección subtítulo */
.seccion-titulo {
    color: #c9d1f5;
    font-size: 1.05rem;
    font-weight: 700;
    margin: 1.5rem 0 0.8rem 0;
    text-transform: uppercase;
    letter-spacing: 1px;
}
</style>
""", unsafe_allow_html=True)

# ── Encabezado ────────────────────────────────────────────────────────────────
st.markdown('<div class="titulo-app">Actualización de Datos COES</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo-app">Consulta el estado de los datos y actualiza la información pendiente del COES.</div>', unsafe_allow_html=True)

# ── Fecha fin seleccionable (por defecto: hoy) ───────────────────────────────
hoy = pd.Timestamp.today().normalize()
fecha_fin_sel = st.date_input(
    "Fecha de actualización",
    value=hoy.date(),
    min_value=FECHA_INICIO_INICIAL.date(),
    max_value=hoy.date(),
)
FECHA_FIN = pd.Timestamp(fecha_fin_sel)

# ── Estado actual de la base ──────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">ESTADO DE LOS DATOS</div>', unsafe_allow_html=True)

TABLAS = ["DEMANDA", MOTIVOS_SHEET, COSTOS_SHEET, INDICES_SHEET, GENERACION_SHEET,
          CMG_SHEET, IEOD_TIPO_SHEET, IEOD_AREAS_SHEET, TERMICAS_SHEET]
# El IEOD (ejecutado) se publica al día siguiente: está "al día" si llega hasta ayer
TABLAS_IEOD = {IEOD_TIPO_SHEET, IEOD_AREAS_SHEET}
DESCRIPCION = {
    "DEMANDA": "Demanda PDO/RDO",
    MOTIVOS_SHEET: "Motivos de reprogramación",
    COSTOS_SHEET: "Costo total",
    INDICES_SHEET: "Índices α β γ",
    GENERACION_SHEET: "Generación por recurso PDO/RDO",
    CMG_SHEET: "CMG programado (3 barras)",
    IEOD_TIPO_SHEET: "IEOD ejecutado por tipo de recurso",
    IEOD_AREAS_SHEET: "IEOD ejecutado por área (Norte/Centro/Sur)",
    TERMICAS_SHEET: "Térmicas gas y diésel por central",
}

@st.cache_data(ttl=60, show_spinner="Consultando fechas en base de datos …")
def _cargar_max_fechas():
    resultado = {}
    for t in TABLAS:
        mx = _max_fecha_pg(t)
        resultado[t] = mx
    return resultado

max_fechas = _cargar_max_fechas()

todo_ok = True
for tabla in TABLAS:
    mx = max_fechas.get(tabla)
    fecha_str = str(mx.date()) if mx else "sin datos"
    objetivo = FECHA_FIN - pd.Timedelta(days=1) if tabla in TABLAS_IEOD else FECHA_FIN
    dias_atraso = (objetivo - mx).days if mx else None
    if mx and mx >= objetivo:
        badge = '<span class="badge-ok">✓ Al día</span>'
    elif mx:
        badge = f'<span class="badge-warn">Pendiente: {dias_atraso} día(s)</span>'
        todo_ok = False
    else:
        badge = '<span class="badge-err">✕ Sin datos</span>'
        todo_ok = False

    st.markdown(f"""
    <div class="card-estado">
        <div>
            <div class="card-tabla">{tabla}</div>
            <div class="card-fecha">{DESCRIPCION.get(tabla, "")} · Último registro: {fecha_str}</div>
        </div>
        {badge}
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if todo_ok:
    st.markdown(f"""
    <div style="background: rgba(18,153,107,0.22); border: 1px solid #6CF5BC;
                border-radius: 12px; padding: 0.9rem 1.2rem; margin-top: 0.5rem;">
        <span style="color: #6CF5BC; font-weight: 600; font-size: 0.95rem;">
            ✅ Todas las tablas están al día hasta <b>{FECHA_FIN.date()}</b>.
        </span>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div style="background: rgba(252,225,34,0.12); border: 1px solid #FCE122;
                border-radius: 12px; padding: 0.9rem 1.2rem; margin-top: 0.5rem;">
        <span style="color: #FCE122; font-weight: 600; font-size: 0.95rem;">
            ⚠️ Hay información pendiente de actualización. Presiona <b>Actualizar datos</b> para completar.
        </span>
    </div>
    """, unsafe_allow_html=True)

# ── Botón de actualización ────────────────────────────────────────────────────
st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
_btn_actualizar = st.button("Actualizar datos", type="primary", use_container_width=True)
_zona_log = st.container()   # aquí sale el avance; el tab 10 se dibuja debajo sin esperar


# ── Tab 10: Potencia Activa (Excel mensual que se sube a mano) ───────────────
st.markdown('<div class="seccion-titulo">POTENCIA ACTIVA (MEDIDORES)</div>', unsafe_allow_html=True)

_MES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun",
              "jul", "ago", "set", "oct", "nov", "dic"]

def _mes_txt(p: pd.Period) -> str:
    return f"{_MES_CORTO[p.month - 1]}-{p.year}"


@st.cache_data(ttl=60, show_spinner=False)
def _dias_por_mes_potencia() -> dict:
    """{'2024-12': días cargados} según POTENCIA_ACTIVA en Supabase."""
    try:
        con = _pg_con()
        cur = con.cursor()
        cur.execute(f'SELECT to_char("Fecha", \'YYYY-MM\'), COUNT(DISTINCT "Fecha") '
                    f'FROM "{POTENCIA_SHEET}" GROUP BY 1')
        res = dict(cur.fetchall())
        con.close()
        return res
    except Exception:
        return {}


# El COES publica el Excel cuando termina el mes: lo esperado llega hasta el mes anterior
_dias_pa = _dias_por_mes_potencia()
_ultimo_esperado = (FECHA_FIN.to_period("M") - 1)
_meses_esp = list(pd.period_range(FECHA_INICIO_INICIAL.to_period("M"), _ultimo_esperado, freq="M"))
_meses_extra = sorted(pd.Period(k, "M") for k in _dias_pa if pd.Period(k, "M") > _ultimo_esperado)

_fichas, _faltan, _incompletos = [], [], []
for _p in _meses_esp + _meses_extra:
    _n = _dias_pa.get(str(_p), 0)
    if _n >= _p.days_in_month:
        _cls = "pa-ok"
    elif _n > 0:
        _cls = "pa-inc"
        _incompletos.append(_p)
    else:
        _cls = "pa-falta"
        _faltan.append(_p)
    _tip = f"{_n} de {_p.days_in_month} días" if _n else "sin cargar"
    _fichas.append(f'<span class="pa-mes {_cls}" title="{_tip}">{_mes_txt(_p)}</span>')

_cargados = [pd.Period(k, "M") for k in _dias_pa]
_ultimo_cargado = max(_cargados) if _cargados else None
_pendientes = len(_faltan) + len(_incompletos)

if not _cargados:
    _badge_pa = '<span class="badge-err">✕ Sin datos</span>'
elif _pendientes == 0:
    _badge_pa = f'<span class="badge-ok">✓ Al día hasta {_mes_txt(_ultimo_esperado)}</span>'
else:
    _badge_pa = f'<span class="badge-warn">Pendiente: {_pendientes} mes(es)</span>'

_hasta_txt = (f"Cargado hasta: <b style='color:#ffffff'>{_mes_txt(_ultimo_cargado)}</b>"
              if _ultimo_cargado else "Sin meses cargados")

st.markdown(f"""
<div class="card-estado card-col">
    <div class="card-cab">
        <div>
            <div class="card-tabla">{POTENCIA_SHEET}</div>
            <div class="card-fecha">Medidores de potencia activa (solares y eólicas) · {_hasta_txt}</div>
        </div>
        {_badge_pa}
    </div>
    <div class="pa-meses">{''.join(_fichas)}</div>
    <div class="pa-leyenda">
        <span class="pa-mes pa-ok">completo</span>&nbsp;
        <span class="pa-mes pa-inc">incompleto</span>&nbsp;
        <span class="pa-mes pa-falta">falta</span>
        &nbsp;· Pasa el mouse sobre un mes para ver cuántos días tiene.
    </div>
</div>
""", unsafe_allow_html=True)

# Mensaje de la última subida (se guarda para mostrarlo después del rerun)
if "pa_resultado" in st.session_state:
    for _l in st.session_state.pop("pa_resultado"):
        _cls = " err" if _l.startswith("✗") else ""
        st.markdown(f'<div class="msg-linea{_cls}">{_l}</div>', unsafe_allow_html=True)

with st.container(border=True):
    st.markdown('<div class="panel-titulo">Subir Excel de Potencia Activa</div>'
                '<div class="panel-ayuda">Archivo mensual del COES «Registros de medidores de generación '
                'cada 15 minutos de potencia activa (MW)». Puedes subir varios meses a la vez; '
                'si un mes ya estaba cargado, se reemplaza.</div>', unsafe_allow_html=True)
    _archivos_pa = st.file_uploader(
        "Excel de medidores", type=["xlsx"], accept_multiple_files=True,
        key="uploader_potencia", label_visibility="collapsed")
    _subir = st.button("Subir Potencia Activa", type="primary", use_container_width=True)
    if _subir and not _archivos_pa:
        st.markdown('<div class="msg-linea err">Primero elige uno o más Excel con «Browse files».</div>',
                    unsafe_allow_html=True)

if _subir and _archivos_pa:
    _lineas, _partes = [], []
    with st.spinner("Leyendo los Excel…"):
        for _a in _archivos_pa:
            try:
                _d = leer_excel_potencia(_a.getvalue(), _a.name)
                _lineas.append(f"✓ {_a.name}: {_d['Fecha'].min()} → {_d['Fecha'].max()} · "
                               f"{_d['Fecha'].nunique()} días · {_d['Central'].nunique()} centrales")
                _partes.append(_d)
            except Exception as _e:
                _lineas.append(f"✗ {_a.name}: {_e}")
    if _partes:
        _df_pa = pd.concat(_partes, ignore_index=True).drop_duplicates(["Fecha", "Central"], keep="last")
        try:
            with st.spinner("Subiendo a Supabase…"):
                _con = _pg_con()
                _n = subir_potencia_pg(_df_pa, _con)
                _con.close()
            _lineas.append(f"✅ {_n:,} filas subidas a POTENCIA_ACTIVA.")
        except Exception as _e:
            _lineas.append(f"✗ No se pudo subir a Supabase: {_e}")
    st.session_state["pa_resultado"] = _lineas
    _dias_por_mes_potencia.clear()
    st.rerun()


# ── Ejecución de la actualización (al final, para que el tab 10 ya esté en pantalla) ──
if _btn_actualizar:
    with _zona_log:
        log_box = st.container()
        def log_msg(msg):
            log_box.markdown(f"""
            <div style="background: rgba(255,255,255,0.10); border-left: 3px solid #6CF5BC;
                        border-radius: 6px; padding: 0.5rem 1rem; margin-bottom: 0.4rem;
                        color: #e0e6ff; font-size: 0.88rem; font-family: monospace;">
                {msg}
            </div>""", unsafe_allow_html=True)

        class _Log:
            def write(self, msg): log_msg(msg)
        log = _Log()
        resumen = {}

        with st.spinner("Conectando a Supabase…"):
            try:
                con = _pg_con()
            except Exception as e:
                st.error(f"No se pudo conectar a Supabase: {e}")
                st.stop()

        # ── DEMANDA ───────────────────────────────────────────────────────────────
        mx_dem = max_fechas.get("DEMANDA")
        inicio_dem = FECHA_INICIO_INICIAL if mx_dem is None else mx_dem   # incluye el último día guardado
        if inicio_dem <= FECHA_FIN:
            df_dem = procesar_demanda_rango(inicio_dem, FECHA_FIN, log)
            n = _insertar_df_pg(df_dem, "DEMANDA", con)
            resumen["DEMANDA"] = n
            log.write(f"  → {n:,} filas insertadas en DEMANDA")
        else:
            log.write("DEMANDA ya está al día ✅")
            resumen["DEMANDA"] = 0

        # ── MOTIVOS / COSTOS / ÍNDICES ────────────────────────────────────────────
        mx_m = max_fechas.get(MOTIVOS_SHEET)
        mx_c = max_fechas.get(COSTOS_SHEET)
        mx_i = max_fechas.get(INDICES_SHEET)
        inicio_m = FECHA_INICIO_INICIAL if mx_m is None else mx_m   # incluye el último día
        inicio_c = FECHA_INICIO_INICIAL if mx_c is None else mx_c
        inicio_i = FECHA_INICIO_INICIAL if mx_i is None else mx_i
        inicio_mci = min(inicio_m, inicio_c, inicio_i)

        if inicio_mci <= FECHA_FIN:
            df_m, df_c, df_i = tab2_procesar_rango(inicio_mci, FECHA_FIN, log)
            nm = _insertar_df_pg(df_m, MOTIVOS_SHEET,  con)
            nc = _insertar_df_pg(df_c, COSTOS_SHEET,   con)
            ni = _insertar_df_pg(df_i, INDICES_SHEET,  con)
            resumen[MOTIVOS_SHEET] = nm
            resumen[COSTOS_SHEET]  = nc
            resumen[INDICES_SHEET] = ni
            log.write(f"  → Motivos: {nm:,} | Costos: {nc:,} | Índices: {ni:,} filas insertadas")
        else:
            log.write("Motivos/Costos/Índices ya están al día ✅")
            resumen[MOTIVOS_SHEET] = resumen[COSTOS_SHEET] = resumen[INDICES_SHEET] = 0

        # ── GENERACION ────────────────────────────────────────────────────────────
        mx_gen = max_fechas.get(GENERACION_SHEET)
        inicio_gen = FECHA_INICIO_INICIAL if mx_gen is None else mx_gen   # incluye el último día
        if inicio_gen <= FECHA_FIN:
            df_gen = tab3_procesar_rango(inicio_gen, FECHA_FIN, log)
            ng = _insertar_df_pg(df_gen, GENERACION_SHEET, con)
            resumen[GENERACION_SHEET] = ng
            log.write(f"  → {ng:,} filas insertadas en GENERACION_RDO")
        else:
            log.write("GENERACION_RDO ya está al día ✅")
            resumen[GENERACION_SHEET] = 0

        # ── CMG (tab 4) ───────────────────────────────────────────────────────────
        mx_cmg = max_fechas.get(CMG_SHEET)
        inicio_cmg = FECHA_INICIO_INICIAL if mx_cmg is None else mx_cmg   # incluye el último día
        if inicio_cmg <= FECHA_FIN:
            df_cmg = tab4_procesar_rango(inicio_cmg, FECHA_FIN, log)
            nc_ = _insertar_df_pg(df_cmg, CMG_SHEET, con)
            resumen[CMG_SHEET] = nc_
            log.write(f"  → {nc_:,} filas insertadas en CMG_RDO")
        else:
            log.write("CMG_RDO ya está al día ✅")
            resumen[CMG_SHEET] = 0

        # ── IEOD (tabs 5 y 6) ─────────────────────────────────────────────────────
        mx_t = max_fechas.get(IEOD_TIPO_SHEET)
        mx_a = max_fechas.get(IEOD_AREAS_SHEET)
        inicio_t = FECHA_INICIO_INICIAL if mx_t is None else mx_t
        inicio_a = FECHA_INICIO_INICIAL if mx_a is None else mx_a
        inicio_ieod = min(inicio_t, inicio_a)          # incluye el último día guardado
        if inicio_ieod <= FECHA_FIN:
            df_t, df_a = tab56_procesar_rango(inicio_ieod, FECHA_FIN, log)
            nt = _insertar_df_pg(df_t, IEOD_TIPO_SHEET, con)
            _actualizar_columnas_g_areas(df_a, con)
            na = _insertar_df_pg(df_a, IEOD_AREAS_SHEET, con)
            resumen[IEOD_TIPO_SHEET] = nt
            resumen[IEOD_AREAS_SHEET] = na
            log.write(f"  → IEOD_TIPO_RECURSO: {nt:,} | IEOD_G_AREAS: {na:,} filas insertadas")
        else:
            log.write("IEOD ya está al día ✅")
            resumen[IEOD_TIPO_SHEET] = resumen[IEOD_AREAS_SHEET] = 0

        # ── TÉRMICAS POR CENTRAL (tabs 7 y 8) ─────────────────────────────────────
        mx_ter = max_fechas.get(TERMICAS_SHEET)
        if mx_ter is None:
            # El histórico se carga con el script termicas_coes.py; aquí solo los últimos días
            inicio_ter = max(FECHA_INICIO_INICIAL, FECHA_FIN - pd.Timedelta(days=2))
            log.write("TERMICAS_RDO está vacía: se cargan solo los últimos 3 días "
                      "(el histórico se sube con termicas_coes.py)")
        else:
            inicio_ter = mx_ter                          # incluye el último día guardado
        if inicio_ter <= FECHA_FIN:
            df_ter = tab78_procesar_rango(inicio_ter, FECHA_FIN, log)
            nter = _insertar_df_pg(df_ter, TERMICAS_SHEET, con)
            resumen[TERMICAS_SHEET] = nter
            log.write(f"  → {nter:,} filas insertadas en TERMICAS_RDO")
        else:
            log.write("TERMICAS_RDO ya está al día ✅")
            resumen[TERMICAS_SHEET] = 0

        con.close()

        # ── Resumen final ─────────────────────────────────────────────────────────
        total_insertadas = sum(resumen.values())
        st.success(f"✅ Actualización completada — {total_insertadas:,} filas nuevas en total.")
        st.cache_data.clear()   # fuerza recarga de fechas al próximo refresh
        st.rerun()
