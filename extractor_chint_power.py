# -*- coding: utf-8 -*-
"""
Extractor de datos - Monitor Chint Power Systems (monitor.chintpowersystems.com)
Listo para Google Colab.

FLUJO:
  1) Login con usuario/clave -> token o cookies de sesion
  2) Modo EXPLORACION: prueba endpoints candidatos y muestra cuales responden
  3) Descarga: lista de plantas, resumen (potencia/energia), dispositivos, historico
  4) Exporta todo a Excel y genera un ZIP para descargar

COMO USAR EN COLAB:
  - Pegar este archivo en una celda o subirlo
  - Completar USUARIO y CLAVE
  - Ejecutar. Revisar el modo exploracion y, si hace falta, ajustar ENDPOINTS.
"""
import requests, json, zipfile, os, re, getpass
import pandas as pd

# ============================= CONFIGURACION =============================
BASE_URL = "https://monitor.chintpowersystems.com"
USUARIO  = ""   # <-- tu usuario del portal
CLAVE    = ""   # <-- tu clave (si lo dejas vacio, te lo pedira al ejecutar)

# Endpoints de la API interna (los usa el navegador; ver modo exploracion).
# Se autodetectan si se dejan en "".
ENDPOINTS = {
    "login":    "",   # ej: "/api/login"  o  "/api/v1/user/login"
    "plants":   "",   # ej: "/api/v1/plant/list"
    "overview": "",   # ej: "/api/v1/plant/overview"  (requiere plant_id)
    "devices":  "",   # ej: "/api/v1/device/list"
    "history":  "",   # ej: "/api/v1/device/history"
}
TOKEN_KEY_CANDIDATES = ["token", "access_token", "accessToken"]  # donde suele venir el token

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json;charset=UTF-8",
    "Referer": BASE_URL + "/",
    "Origin": BASE_URL,
}

# ============================ CANDIDATOS (exploracion) ============================
CANDIDATE_PATHS = [
    "/api/login", "/api/user/login", "/api/v1/login", "/api/v1/user/login",
    "/api/auth/login", "/api/v1/auth/login",
    "/api/plant/list", "/api/plant/getPlantList", "/api/v1/plant/list",
    "/api/v1/plants", "/api/plant", "/api/v1/plant/page",
    "/api/device/list", "/api/v1/device/list",
    "/api/plant/overview", "/api/v1/plant/overview",
    "/api/v1/plant/summary", "/api/plant/summary",
    "/api/station/list", "/api/v1/station/list",
    "/api/user/info", "/api/v1/user/info",
]
LOGIN_PAYLOAD_SHAPES = [
    lambda u, p: {"username": u, "password": p},
    lambda u, p: {"userName": u, "password": p},
    lambda u, p: {"account": u, "password": p},
    lambda u, p: {"email": u, "password": p},
]

session = requests.Session()
session.headers.update(HEADERS)
TOKEN = None


# ================================ FUNCIONES ================================
def _url(path):
    return BASE_URL + path if path.startswith("/") else path


def _extraer_token(data):
    """Busca el token en distintas profundidades del JSON de respuesta."""
    if not isinstance(data, dict):
        return None
    for k in TOKEN_KEY_CANDIDATES:
        if data.get(k):
            return data[k]
    for sub in ("data", "result", "body"):
        if isinstance(data.get(sub), dict):
            t = _extraer_token(data[sub])
            if t:
                return t
    return None


def login():
    global TOKEN
    if not USUARIO:
        u = input("Usuario del portal: ").strip()
    else:
        u = USUARIO
    p = CLAVE or getpass.getpass("Clave del portal: ")

    tried = set()
    for shape in LOGIN_PAYLOAD_SHAPES:
        paths = [ENDPOINTS["login"]] if ENDPOINTS["login"] else \
                [c for c in CANDIDATE_PATHS if "login" in c]
        for path in paths:
            if path in tried:
                continue
            tried.add(path)
            try:
                r = session.post(_url(path), data=json.dumps(shape(u, p)), timeout=30)
            except Exception as e:
                print(f"  [login] {path} -> ERROR de red: {e}")
                continue
            print(f"  [login] {path} -> HTTP {r.status_code}")
            try:
                data = r.json()
            except Exception:
                continue
            if r.status_code == 200 and data:
                token = _extraer_token(data)
                if token:
                    TOKEN = token
                    session.headers["Authorization"] = f"Bearer {TOKEN}"
                    ENDPOINTS["login"] = path
                    print(f"  >>> LOGIN OK en {path} (token capturado)")
                    return True
                # algunos portales usan solo cookies de sesion
                if session.cookies.get_dict():
                    ENDPOINTS["login"] = path
                    print(f"  >>> LOGIN OK en {path} (sesion por cookies)")
                    return True
    print("No se pudo iniciar sesion automaticamente. Revisa credenciales o captura")
    print("el endpoint real con F12 -> Network en el navegador y ponlo en ENDPOINTS['login'].")
    return False


def explorar():
    """Prueba endpoints GET candidatos y reporta cuales responden con JSON."""
    print("\n--- MODO EXPLORACION (GET sobre endpoints candidatos) ---")
    for path in CANDIDATE_PATHS:
        if "login" in path:
            continue
        try:
            r = session.get(_url(path), timeout=20,
                            params={"page": 1, "size": 10} if "list" in path or "page" in path else None)
        except Exception as e:
            print(f"  {path} -> ERROR: {e}")
            continue
        marca = ""
        try:
            data = r.json()
            keys = list(data.keys())[:6] if isinstance(data, dict) else f"lista[{len(data)}]"
            marca = f" JSON keys={keys}"
            if r.status_code == 200 and data not in ({}, [], None):
                print(f"  *** {path} -> HTTP {r.status_code}{marca}  <-- CANDIDATO FUERTE")
        except Exception:
            marca = f" (no JSON: {r.text[:60]!r})"
        if r.status_code != 200 or not marca.startswith(" JSON"):
            print(f"  {path} -> HTTP {r.status_code}{marca}")
    print("Tip: si ves rutas que funcionan, copialas en ENDPOINTS y vuelve a ejecutar.\n")


def _resolver(nombre, candidatos):
    if ENDPOINTS[nombre]:
        return ENDPOINTS[nombre]
    for c in candidatos:
        try:
            r = session.get(_url(c), params={"page": 1, "size": 5}, timeout=20)
            if r.status_code == 200:
                try:
                    if r.json():
                        ENDPOINTS[nombre] = c
                        return c
                except Exception:
                    pass
        except Exception:
            pass
    return None


def _a_dataframe(data):
    """Convierte una respuesta JSON arbitraria a DataFrame lo mejor posible."""
    if isinstance(data, dict):
        for k in ("data", "result", "rows", "list", "records"):
            if isinstance(data.get(k), list) and data[k]:
                return pd.json_normalize(data[k])
        for k in ("data", "result"):
            if isinstance(data.get(k), dict):
                return _a_dataframe(data[k])
        return pd.json_normalize(data)
    if isinstance(data, list):
        return pd.json_normalize(data)
    return pd.DataFrame({"respuesta": [str(data)[:500]]})


def descargar_datos():
    frames = {}

    ep = _resolver("plants", [c for c in CANDIDATE_PATHS if "plant" in c and "list" in c or "station" in c])
    if not ep:
        print("No se encontro endpoint de plantas. Usa el modo exploracion.")
        return None
    r = session.get(_url(ep), params={"page": 1, "size": 500}, timeout=30)
    plants = r.json()
    df_plants = _a_dataframe(plants)
    frames["Plantas"] = df_plants
    print(f"Plantas encontradas: {len(df_plants)} (endpoint {ep})")

    plant_ids = []
    for col in df_plants.columns:
        if col.lower() in ("plant_id", "plantid", "id", "station_id", "stationid"):
            plant_ids = df_plants[col].dropna().tolist()
            break
    if not plant_ids and len(df_plants):
        plant_ids = [None]  # el endpoint de resumen a veces no necesita id

    for pid in plant_ids[:20]:
        label = f"Resumen_{pid}" if pid else "Resumen"
        ep2 = _resolver("overview", [c for c in CANDIDATE_PATHS if "overview" in c or "summary" in c])
        if ep2:
            try:
                r2 = session.get(_url(ep2), params={"plant_id": pid} if pid else {}, timeout=30)
                frames[label] = _a_dataframe(r2.json())
            except Exception as e:
                print(f"  aviso: no se pudo leer resumen de {pid}: {e}")

    ep3 = _resolver("devices", [c for c in CANDIDATE_PATHS if "device" in c])
    if ep3:
        r3 = session.get(_url(ep3), params={"page": 1, "size": 500}, timeout=30)
        frames["Dispositivos"] = _a_dataframe(r3.json())

    # Historico por dispositivo (dia/mes/ano) - estructura tipica de portales solares
    ep4 = _resolver("history", [c for c in CANDIDATE_PATHS if "history" in c or "energy" in c or "chart" in c])
    if ep4 and "Dispositivos" in frames:
        df_dev = frames["Dispositivos"]
        sn_col = next((c for c in df_dev.columns if "sn" in c.lower() or "serial" in c.lower()), None)
        if sn_col:
            for sn in df_dev[sn_col].dropna().head(10):
                for tipo in ("day", "month", "year"):
                    try:
                        r4 = session.get(_url(ep4), params={"sn": sn, "type": tipo}, timeout=30)
                        frames[f"Hist_{str(sn)[:12]}_{tipo}"] = _a_dataframe(r4.json())
                    except Exception:
                        pass
    return frames


def exportar(frames):
    if not frames:
        print("Sin datos para exportar.")
        return
    xlsx = "/mnt/data/chint_power_data.xlsx" if os.path.exists("/mnt/data") else "chint_power_data.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        for nombre, df in frames.items():
            df.to_excel(w, sheet_name=re.sub(r"[^\w]", "_", nombre)[:31], index=False)
    zipp = xlsx.replace(".xlsx", ".zip")
    with zipfile.ZipFile(zipp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(xlsx, os.path.basename(xlsx))
    print(f"\nOK: {xlsx}  +  {zipp}")
    try:
        from google.colab import files
        files.download(zipp)   # descarga automatica del ZIP en Colab
        print("Descarga del ZIP iniciada.")
    except ImportError:
        print("No estas en Colab: toma los archivos del directorio actual.")


# ================================= MAIN =================================
if __name__ == "__main__":
    if login():
        if not all(ENDPOINTS.values()):
            explorar()
        datos = descargar_datos()
        exportar(datos)
