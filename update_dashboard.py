#!/usr/bin/env python3
"""
Robot de actualización del Tablero Ejecutivo — Agricola Guapa
----------------------------------------------------------------
Se ejecuta dentro de GitHub Actions. Hace 4 cosas:
  1. Se autentica ante Microsoft Graph con las credenciales de la app
     (usuario técnico, sin persona humana conectada).
  2. Busca y descarga el Excel más reciente desde el OneDrive indicado.
  3. Extrae y transforma los datos (misma lógica que usamos en el chat).
  4. Rellena la plantilla del tablero (dashboard_template.html) con los
     datos frescos y escribe index.html — el archivo que GitHub Pages sirve.

Variables de entorno esperadas (se configuran como "Secrets" en GitHub,
nunca quedan escritas en este archivo):
  AZURE_TENANT_ID       - Id. de directorio (inquilino)
  AZURE_CLIENT_ID       - Id. de aplicación (cliente)
  AZURE_CLIENT_SECRET   - El "Valor" del secreto de cliente
  ONEDRIVE_USER_EMAIL   - Correo del OneDrive donde vive el Excel
  EXCEL_FILE_NAME       - Nombre (o parte del nombre) del archivo a buscar
  EXCEL_FOLDER_PATHS    - (opcional) lista de carpetas donde buscar primero,
                           separadas por "|". Ej: "|Planta" busca primero en
                           la raíz y luego dentro de la carpeta "Planta".
"""
import os
import re
import sys
import json
import time
import datetime
import unicodedata
import requests
import openpyxl
from io import BytesIO
from urllib.parse import quote

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USER_EMAIL = os.environ["ONEDRIVE_USER_EMAIL"]
FILE_NAME = os.environ.get("EXCEL_FILE_NAME", "Indicadores de empacadora")
FOLDER_PATHS = os.environ.get("EXCEL_FOLDER_PATHS", "|Planta").split("|")

GRAPH = "https://graph.microsoft.com/v1.0"


def log(msg):
    print(f"[robot] {msg}", flush=True)


def get_access_token():
    log("Solicitando token de acceso a Microsoft...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data, timeout=30)
    if r.status_code != 200:
        log(f"ERROR al obtener token: {r.status_code} {r.text}")
        sys.exit(1)
    log("Token obtenido correctamente.")
    return r.json()["access_token"]


def find_excel_file(token):
    headers = {"Authorization": f"Bearer {token}"}

    # 1) intento directo: probamos el archivo en cada carpeta candidata
    #    (por defecto: la raíz, y luego la carpeta "Planta") — es más rápido
    #    y confiable que depender del buscador de Microsoft.
    for folder in FOLDER_PATHS:
        for ext in (".xlsx", ".xlsm"):
            candidate = f"{FILE_NAME}{ext}"
            path = f"{folder}/{candidate}" if folder else candidate
            ubicacion = f"carpeta '{folder}'" if folder else "la raíz"
            log(f"Intentando acceso directo: '{candidate}' en {ubicacion} del OneDrive de {USER_EMAIL}...")
            url = f"{GRAPH}/users/{USER_EMAIL}/drive/root:/{quote(path)}"
            r = requests.get(url, headers=headers, timeout=30)
            if r.status_code == 200:
                item = r.json()
                log(f"Encontrado directamente: '{item['name']}' en {ubicacion} (modificado {item.get('lastModifiedDateTime')})")
                return item["id"]
            log(f"  no encontrado así ({r.status_code}), probando siguiente opción...")

    # 2) respaldo: buscarlo por nombre en todo el OneDrive, por si está en
    #    otra carpeta que no probamos arriba, o cambió un poco el nombre.
    #    Reintentamos unas veces porque este buscador puede fallar de forma
    #    pasajera (p.ej. si OneDrive está sincronizando el archivo justo en
    #    ese momento).
    last_error = None
    for intento in range(1, 4):
        log(f"Buscando '{FILE_NAME}' en todo el OneDrive de {USER_EMAIL}... (intento {intento}/3)")
        url = f"{GRAPH}/users/{USER_EMAIL}/drive/root/search(q='{quote(FILE_NAME)}')"
        r = requests.get(url, headers=headers, timeout=30)
        if r.status_code == 200:
            items = r.json().get("value", [])
            items = [i for i in items if i.get("name", "").lower().endswith((".xlsx", ".xlsm"))]
            if items:
                items.sort(key=lambda i: i.get("lastModifiedDateTime", ""), reverse=True)
                chosen = items[0]
                log(f"Encontrado por búsqueda: '{chosen['name']}' (modificado {chosen.get('lastModifiedDateTime')})")
                return chosen["id"]
            log("  la búsqueda no encontró ningún .xlsx/.xlsm que coincida.")
            last_error = "sin resultados"
        else:
            log(f"  ERROR al buscar el archivo: {r.status_code} {r.text}")
            last_error = f"{r.status_code} {r.text}"
        if intento < 3:
            log("  esperando 10 segundos antes de reintentar...")
            time.sleep(10)

    log(f"ERROR: no encontré el archivo, ni directo ni buscando, tras 3 intentos. Último error: {last_error}")
    sys.exit(1)


def download_file(token, item_id):
    log("Descargando el archivo...")
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH}/users/{USER_EMAIL}/drive/items/{item_id}/content"
    r = requests.get(url, headers=headers, timeout=60)
    if r.status_code != 200:
        log(f"ERROR al descargar: {r.status_code} {r.text}")
        sys.exit(1)
    log(f"Descargado ({len(r.content)} bytes).")
    return BytesIO(r.content)


# ================= misma lógica de extracción que ya usamos en el chat =================

def is_err(v):
    return isinstance(v, str) and v.startswith("#")


def num(v, nd=3):
    if is_err(v) or v is None:
        return 0
    if isinstance(v, (int, float)):
        return round(v, nd)
    return v


def time_to_hours(t):
    if t is None:
        return 0
    if isinstance(t, datetime.time):
        return round(t.hour + t.minute / 60 + t.second / 3600, 3)
    if isinstance(t, (int, float)):
        return round(t, 3)
    return 0


# ================= NÓMINA DIARIA + PARÁMETROS =================

def _norm(s):
    """minúsculas, sin tildes y con espacios simples (para comparar encabezados)"""
    s = unicodedata.normalize("NFKD", str(s if s is not None else "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _find_sheet(wb, wanted):
    for name in wb.sheetnames:
        if _norm(name) == wanted:
            return wb[name]
    return None


def _to_float(v):
    """número desde celda: vacío / espacio / texto raro -> 0; admite coma decimal"""
    if v is None or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


NOMINA_COLS = {
    "fecha": "fecha", "cedula": "cedula", "nombre completo": "nombre",
    "codigo labor": "cod", "horas ordinarias": "ord",
    "rn": "rn", "hed": "hed", "hen": "hen", "rfn": "rfn", "rfd": "rfd",
    "fn": "fn", "fd": "fd", "hedf": "hedf", "henf": "henf",
}
NOMINA_FIELDS = ["fecha", "cedula", "cod", "ord", "rn", "hed", "hen", "rfn", "rfd", "fn", "fd", "hedf", "henf"]


def extract_nomina(wb):
    """Lee la hoja 'Nomina diaria' (reportes semanales pegados uno debajo del otro).
    - encuentra el encabezado por nombre de columna (no por posición)
    - salta filas vacías y encabezados repetidos
    - descarta solo filas IDÉNTICAS (cédula+fecha+hora inicio+hora final+código):
      una misma persona puede tener 2 filas el mismo día (p.ej. turno + capacitación)
    """
    vacio = {"fields": NOMINA_FIELDS, "rows": []}
    ws = _find_sheet(wb, "nomina diaria")
    if ws is None:
        log('AVISO: no encontré la hoja "Nomina diaria". El panel de costos de nómina quedará vacío.')
        return vacio, {}
    colmap, rows, personas, vistos = {}, [], {}, set()
    descartadas = {"sin_fecha": 0, "sin_cedula": 0, "sin_codigo": 0, "repetidas": 0}
    for r in ws.iter_rows(values_only=True):
        celdas = [_norm(c) for c in r]
        if "fecha" in celdas and "codigo labor" in celdas:      # encabezado (o encabezado repetido)
            colmap = {}
            for idx, c in enumerate(celdas):
                if c in NOMINA_COLS:
                    colmap[NOMINA_COLS[c]] = idx
                elif c == "hora inicio":
                    colmap["h_ini"] = idx
                elif c == "hora final":
                    colmap["h_fin"] = idx
            continue
        if not colmap:
            continue
        get = lambda k: r[colmap[k]] if k in colmap and colmap[k] < len(r) else None
        if all(c is None or (isinstance(c, str) and not c.strip()) for c in r):
            continue
        fecha = _to_date(get("fecha"))
        if fecha is None:
            if _to_float(get("cedula")):          # tiene cédula pero fecha ilegible: avisar
                descartadas["sin_fecha"] += 1
            continue                              # (filas tipo "JORNAL" o de relleno: se saltan en silencio)
        cedula = int(_to_float(get("cedula")))
        if not cedula:
            descartadas["sin_cedula"] += 1
            continue
        cod = int(_to_float(get("cod")))
        if not cod:
            descartadas["sin_codigo"] += 1
            continue
        clave = (cedula, fecha, str(get("h_ini")).strip(), str(get("h_fin")).strip(), cod)
        if clave in vistos:
            descartadas["repetidas"] += 1
            continue
        vistos.add(clave)
        nombre = re.sub(r"\s+", " ", str(get("nombre") or "")).strip()
        if nombre:
            personas[str(cedula)] = nombre
        horas = [round(_to_float(get(k)), 2) for k in NOMINA_FIELDS[3:]]
        rows.append([fecha.isoformat(), cedula, cod] + horas)
    rows.sort(key=lambda x: (x[0], x[1], x[2]))
    log(f"Nomina diaria: {len(rows)} registros, {len(personas)} personas"
        + (f", {rows[0][0]} a {rows[-1][0]}" if rows else "")
        + f" | descartadas: {descartadas}")
    return {"fields": NOMINA_FIELDS, "rows": rows}, personas


def _as_pct(v):
    x = _to_float(v)
    return x / 100.0 if x > 1.5 else x


def extract_params_nomina(wb):
    """Lee la hoja 'Parametros nomina' por títulos de sección. Si falta algo, el
    tablero usa valores por defecto (ver DEFAULT_PARAMS en la plantilla)."""
    params = {"costos": [], "vigencias": [], "reglas": {}}
    ws = _find_sheet(wb, "parametros nomina")
    if ws is None:
        log('AVISO: no encontré la hoja "Parametros nomina"; se usarán los valores por defecto.')
        return params
    filas = [list(r) for r in ws.iter_rows(values_only=True)]
    seccion = None
    for r in filas:
        a = r[0] if r else None
        na = _norm(a)
        if na.startswith("costo por dia por codigo"):
            seccion = "costos"; continue
        if na.startswith("tarifas de horas extra"):
            seccion = "vigencias"; continue
        if na.startswith("reglas de calculo"):
            seccion = "reglas"; continue
        if na.startswith("referencia"):
            seccion = None; continue
        if seccion == "costos":
            if isinstance(a, (int, float)) and not isinstance(a, bool):
                costo = r[2] if len(r) > 2 else None
                entra = _norm(r[3] if len(r) > 3 else "") == "si"
                params["costos"].append({
                    "cod": int(a), "grupo": str(r[1] or "").strip(),
                    "costo": (None if costo is None or str(costo).strip() == "" else _to_float(costo)),
                    "entra": entra})
        elif seccion == "vigencias":
            desde = _to_date(a)
            if desde:
                hasta = _to_date(r[1] if len(r) > 1 else None)
                params["vigencias"].append({
                    "desde": desde.isoformat(), "hasta": hasta.isoformat() if hasta else None,
                    "salario": _to_float(r[2] if len(r) > 2 else 0),
                    "divisor": _to_float(r[3] if len(r) > 3 else 0),
                    "dominical": _as_pct(r[4] if len(r) > 4 else 0)})
        elif seccion == "reglas":
            if na and len(r) > 1 and r[1] is not None and na != "regla":
                v = r[1]
                if "domingos y festivos" in na:
                    params["reglas"]["cobrar_base_domingo"] = _norm(v) == "si"
                elif "recargos" in na:
                    params["reglas"]["contar_recargos"] = _norm(v) == "si"
                elif na.startswith("recargo nocturno"):
                    params["reglas"]["rn_pct"] = _as_pct(v)
                elif "extra diurna" in na:
                    params["reglas"]["hed_pct"] = _as_pct(v)
                elif "extra nocturna" in na:
                    params["reglas"]["hen_pct"] = _as_pct(v)
    log(f"Parametros nomina: {len(params['costos'])} códigos, {len(params['vigencias'])} vigencias, reglas={params['reglas']}")
    return params


def extract_data(file_bytes):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    START_DATE = datetime.datetime(2023, 1, 2)

    ws0 = wb["Indicadores de planta"]
    last_real_date = None
    for r in ws0.iter_rows(min_row=3, values_only=True):
        fecha = r[0]
        if not isinstance(fecha, datetime.datetime):
            continue
        kg = r[4]
        if isinstance(kg, (int, float)) and kg > 0:
            if last_real_date is None or fecha > last_real_date:
                last_real_date = fecha
    END_DATE = last_real_date or datetime.datetime.now()
    log(f"Última fecha con producción real detectada: {END_DATE.date()}")

    # 1) Indicadores de planta (diario)
    ws = wb["Indicadores de planta"]
    fields1 = ["fecha", "semana", "mes", "anio", "kg_dia", "bines", "peso_bin", "bin_hora",
               "kg_hora", "pinas", "ocupacion", "rendimiento_planta", "descarte",
               "eficiencia_planta", "eficiencia_mo", "rendimiento_mo", "operarios",
               "kg_hombre", "uni_hombre"]
    rows1 = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        fecha = r[0]
        if not isinstance(fecha, datetime.datetime):
            continue
        if fecha < START_DATE or fecha > END_DATE:
            continue
        kg = r[4]
        if is_err(kg):
            kg = 0
        rows1.append([
            fecha.strftime("%Y-%m-%d"), int(r[1]) if r[1] else 0, r[2], int(r[3]) if r[3] else 0,
            num(kg, 0), num(r[5], 0), num(r[6], 1), num(r[7], 2), num(r[8], 1), num(r[9], 0),
            num(r[12], 3), num(r[13], 4), num(r[14], 4), num(r[15], 4), num(r[16], 4),
            num(r[17], 4), num(r[18], 0), num(r[19], 1), num(r[20], 1),
        ])
    log(f"Indicadores de planta: {len(rows1)} filas")

    # 2) Horas extras (todos los años disponibles)
    fields2 = ["anio", "semana", "horas_extras_total", "kilos_procesados", "recargo_nocturno",
               "horas_extra_ordinarias", "recargo_festivo", "horas_extra_festivas"]
    rows2 = []
    for name in wb.sheetnames:
        if not name.strip().lower().startswith("horas extras"):
            continue
        import re
        m = re.search(r"(\d{4})", name)
        if not m:
            continue
        year = int(m.group(1))
        ws = wb[name]
        header = [c.value if hasattr(c, "value") else c for c in next(ws.iter_rows(min_row=1, max_row=1))]
        has_full_cols = len(header) > 6
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[1] is None:
                continue
            if has_full_cols and len(r) > 7:
                rows2.append([year, int(r[1]), num(r[2], 1), num(r[3], 0), num(r[4], 1),
                              num(r[5], 1), num(r[6], 1), num(r[7], 1)])
            else:
                rows2.append([year, int(r[1]), num(r[2], 1), num(r[3], 0), 0,
                              num(r[4], 1), 0, num(r[5], 1)])
    log(f"Horas extras: {len(rows2)} filas")

    # 3) Tiempomuertoanual — detección robusta de columnas (soporta causas nuevas)
    ws = wb["Tiempomuertoanual"]
    headers = [c for c in ws.iter_rows(min_row=1, max_row=1, values_only=True)][0]
    tm_total_idx = None
    horas_trab_idx = None
    for idx, hd in enumerate(headers):
        hdn = str(hd).lower() if hd else ""
        if "tiempo muerto" in hdn and tm_total_idx is None:
            tm_total_idx = idx
        if "horas trabajadas" in hdn:
            horas_trab_idx = idx
    causa_names = [headers[i].strip() for i in range(2, tm_total_idx)]
    fields3 = ["fecha", "semana"] + causa_names + ["tiempo_muerto_total_hr", "horas_trabajadas"]
    rows3 = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        fecha = r[0]
        if not isinstance(fecha, datetime.datetime):
            continue
        if fecha < START_DATE or fecha > END_DATE:
            continue
        causas_vals = [time_to_hours(r[i]) if i < len(r) else 0 for i in range(2, tm_total_idx)]
        tm_total = time_to_hours(r[tm_total_idx]) if tm_total_idx < len(r) else 0
        ht = time_to_hours(r[horas_trab_idx]) if horas_trab_idx and horas_trab_idx < len(r) else 0
        rows3.append([fecha.strftime("%Y-%m-%d"), int(r[1]) if r[1] else 0] + causas_vals + [tm_total, ht])
    log(f"Tiempomuertoanual: {len(rows3)} filas, {len(causa_names)} causas ({causa_names})")

    # 4) Consolidado Post-cosecha
    ws = wb["Consolidado Post-cosecha"]
    headers = [c for c in ws.iter_rows(min_row=1, max_row=1, values_only=True)][0]
    causa_names_pc = [headers[i].strip() for i in range(1, 12)]
    fields4 = ["fecha"] + causa_names_pc + ["total_merma", "total_cosecha", "total_despacho"]
    rows4 = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        fecha = r[0]
        if not isinstance(fecha, datetime.datetime):
            continue
        if fecha < START_DATE or fecha > END_DATE:
            continue
        causas_vals = [num(r[i], 0) if i < len(r) else 0 for i in range(1, 12)]
        rows4.append([fecha.strftime("%Y-%m-%d")] + causas_vals + [num(r[12], 0), num(r[13], 0), num(r[14], 0)])
    log(f"Consolidado Post-cosecha: {len(rows4)} filas")

    # 5) TiempoE (fuente exacta de "Datos Eduar")
    ws = wb["TiempoE"]
    rowsE = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        semana, anio, kg, horas_totales, tiempo_muerto, horas_netas = r[0], r[1], r[2], r[3], r[4], r[5]
        if not isinstance(semana, (int, float)) or not isinstance(anio, (int, float)):
            continue
        if anio not in (2023, 2024, 2025, 2026, 2027):
            continue
        rowsE.append([int(anio), int(semana), kg or 0, horas_totales or 0, tiempo_muerto or 0, horas_netas or 0])
    while rowsE and rowsE[-1][5] == 0:
        rowsE.pop()
    log(f"TiempoE: {len(rowsE)} filas")

    nomina, personas = extract_nomina(wb)
    params_nomina = extract_params_nomina(wb)

    return {
        "nomina": nomina,
        "personas": personas,
        "params_nomina": params_nomina,
        "daily": {"fields": fields1, "rows": rows1},
        "horas_extras": {"fields": fields2, "rows": rows2},
        "tiempo_muerto": {"fields": fields3, "rows": rows3},
        "post_cosecha": {"fields": fields4, "rows": rows4},
        "tiempo_e": {"fields": ["anio", "semana", "kg", "horas_totales", "tiempo_muerto", "horas_netas"], "rows": rowsE},
    }


def main():
    token = get_access_token()
    item_id = find_excel_file(token)
    file_bytes = download_file(token, item_id)

    log("Extrayendo datos del Excel...")
    data = extract_data(file_bytes)

    log("Rellenando la plantilla del tablero...")
    with open("dashboard_template.html", "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("__DATA_PLACEHOLDER__", json.dumps(data, ensure_ascii=False, separators=(",", ":")))

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    log(f"index.html generado correctamente ({len(html)} caracteres).")


if __name__ == "__main__":
    main()
