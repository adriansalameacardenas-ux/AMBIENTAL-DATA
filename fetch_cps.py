# -*- coding: utf-8 -*-
"""
fetch_cps.py — Descarga datos del portal CPS y guarda data.json
Corre en GitHub Actions (o local/Colab). Credenciales por variables de entorno
(o GitHub Secrets): CPS_USER, CPS_PASS.
"""
import os, json, requests, datetime

BASE  = os.environ.get("CPS_BASE", "https://monitor.chintpowersystems.com")
USER  = os.environ.get("CPS_USER", "")
PASSW = os.environ.get("CPS_PASS", "")
EP = {  # opcionales: si se definen, se usan directo; si no, autodetección
    "login":   os.environ.get("CPS_LOGIN", ""),
    "plants":  os.environ.get("CPS_PLANTS", ""),
    "devices": os.environ.get("CPS_DEVICES", ""),
    "history": os.environ.get("CPS_HISTORY", ""),
}
LOGIN_PATHS  = ["/api/login", "/api/user/login", "/api/v1/login", "/api/v1/user/login",
                "/api/auth/login", "/api/v1/auth/login", "/login"]
PLANT_PATHS  = ["/api/plant/list", "/api/plant/getPlantList", "/api/v1/plant/list",
                "/api/v1/plants", "/api/plant", "/api/v1/plant/page",
                "/api/station/list", "/api/v1/station/list"]
DEVICE_PATHS = ["/api/device/list", "/api/v1/device/list", "/api/device", "/api/v1/device/page"]
HIST_PATHS   = ["/api/device/history", "/api/v1/device/history", "/api/plant/energy",
                "/api/v1/plant/energy", "/api/v1/energy", "/api/plant/chart"]

s = requests.Session()
s.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0",
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json;charset=UTF-8",
    "Referer": BASE + "/", "Origin": BASE,
})
TOKEN = None
sources = []

def url(p): return p if p.startswith("http") else BASE + p

def buscar_token(d):
    if not isinstance(d, dict): return None
    for k in ("token", "access_token", "accessToken", "jwt"):
        if d.get(k): return d[k]
    for k in ("data", "result", "body"):
        t = buscar_token(d.get(k))
        if t: return t
    return None

def login():
    global TOKEN
    formas = [{"username": USER, "password": PASSW}, {"userName": USER, "password": PASSW},
              {"account": USER, "password": PASSW}, {"email": USER, "password": PASSW}]
    paths = [EP["login"]] if EP["login"] else LOGIN_PATHS
    for p in paths:
        for body in formas:
            try:
                r = s.post(url(p), data=json.dumps(body), timeout=30)
            except Exception:
                continue
            if r.status_code != 200:
                continue
            try:
                data = r.json()
            except Exception:
                continue
            t = buscar_token(data)
            if t:
                TOKEN = t
                s.headers["Authorization"] = "Bearer " + t
                print("Login OK en", p, "(token)")
                return True
            if s.cookies.get_dict():
                print("Login OK en", p, "(cookies)")
                return True
    raise SystemExit("ERROR: login falló. Define CPS_LOGIN con el endpoint real "
                     "(F12 -> Network al iniciar sesión en el portal).")

def traer(paths, params=None):
    for p in paths:
        try:
            r = s.get(url(p), params=params or {"page": 1, "size": 500}, timeout=30)
            if r.status_code == 200:
                data = r.json()
                if data not in ({}, [], None):
                    print("OK:", p)
                    return p, data
        except Exception:
            pass
    return None, None

login()

for nombre, paths, ep_key in [
    ("plantas", PLANT_PATHS, "plants"), ("dispositivos", DEVICE_PATHS, "devices"),
    ("historico", HIST_PATHS, "history")]:
    lista = [EP[ep_key]] if EP[ep_key] else paths
    p, data = traer(lista)
    if data is not None:
        sources.append({"name": nombre, "json": data})

out = {
    "fetched_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    "sources": sources,
}
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False)
print("data.json guardado con", len(sources), "fuentes")
