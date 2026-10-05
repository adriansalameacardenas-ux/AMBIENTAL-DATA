# ☀️ Extractor Chint Power Systems → Excel

Herramienta **no oficial** para extraer datos del portal de monitoreo
[monitor.chintpowersystems.com](https://monitor.chintpowersystems.com) (plantas solares,
inversores, históricos de energía) y exportarlos organizados a **Excel (.xlsx)** o **ZIP**.

Incluye dos formas de uso:

| Archivo | Uso |
|---|---|
| `index.html` | Herramienta 100 % en el navegador (doble clic, sin instalar nada) |
| `extractor_chint_power.py` | Script para **Google Colab** (Python), descarga automática del ZIP |

---

## 🚀 Uso rápido (HTML)

1. Descarga `index.html` y ábrelo con doble clic (Chrome, Edge o Firefox).
2. **Opción A – API directa:** completa URL, usuario, clave y los endpoints que captures
   con `F12 → Network → Fetch/XHR` al iniciar sesión en el portal, y pulsa
   *Conectar y traer datos*.
3. **Opción B – Pegar JSON (recomendada, no depende de CORS):** en el portal, con F12 abierto,
   copia las respuestas de las peticiones (`Copy response`) y péguelas en las cajas
   correspondientes (hay botones 📥 que leen el portapapeles). Acepta varios JSON de
   histórico separados por línea en blanco.
4. Revisa la vista previa y descarga el `.xlsx` multi-hoja o el `.zip`.

## 🚀 Uso rápido (Colab)

1. Abre [colab.research.google.com](https://colab.research.google.com), crea un notebook.
2. Sube `extractor_chint_power.py` o pégalo en una celda con `%run`.
3. Completa `USUARIO` (la clave te la pide al ejecutar).
4. Ejecuta: hace login, **detecta automáticamente los endpoints** (modo exploración),
   descarga plantas / dispositivos / históricos y genera el ZIP con descarga automática.

## 🔎 Cómo capturar los endpoints reales

El portal no documenta su API pública, así que:

1. Inicia sesión en el portal con `F12 → Network → Fetch/XHR` abierto.
2. Filtra por `plant`, `device`, `energy` o `history`.
3. Copia la **URL** (ruta después del dominio) y el **payload** del request.
4. En el HTML, pega esas rutas en los campos de endpoint; en Colab, en el dict `ENDPOINTS`.

## ⚠️ Aviso legal

- Proyecto **independiente y no oficial**: no está afiliado ni respaldado por
  Chint Power Systems / CPS.
- Usa **únicamente tus propias credenciales** y datos de plantas a las que tengas acceso.
- Revisa los [Términos de Servicio](https://monitor.chintpowersystems.com) del portal
  antes de automatizar descargas; respeta límites de peticiones.
- Las credenciales se procesan localmente (HTML) o en tu propia sesión de Colab;
  este código no envía datos a terceros.

## 📦 Estructura

```
extractor-chint-power/
├── index.html              # herramienta web (SheetJS + JSZip por CDN)
├── extractor_chint_power.py# script para Google Colab
├── LICENSE                 # MIT
└── README.md
```

## 🛠️ Tecnologías

- [SheetJS](https://sheetjs.com/) — generación de `.xlsx` en el navegador.
- [JSZip](https://stuk.github.io/jszip/) — compresión a `.zip` en el navegador.
- Python + `requests` + `pandas` + `openpyxl` (Colab).

## 📄 Licencia

MIT — ver [LICENSE](LICENSE).

## 🔄 Auto-actualización (GitHub Actions)

El dashboard se actualiza solo: una Action corre `fetch_cps.py` cada 30 minutos,
guarda `data.json` en el repo y el dashboard lo lee al abrirse (GitHub Pages).

**Configuración (una sola vez):**
1. Repo → **Settings → Secrets and variables → Actions → New repository secret**
2. Crea `CPS_USER` (tu usuario) y `CPS_PASS` (tu clave). Jamás se muestran ni salen del repo.
3. Repo → **Settings → Pages → Source: GitHub Actions** (o rama `main`, carpeta raíz).
4. Abre la pestaña **Actions** → *Actualizar datos CPS* → **Run workflow** para probar.

Si el login automático no encuentra los endpoints, crea los secrets `CPS_LOGIN`,
`CPS_PLANTS`, `CPS_DEVICES` con las rutas capturadas en F12 → Network.
Nota: `data.json` queda público si el repo es público (son datos de energía;
si prefieres mantenerlos privados, deja el repo privado y descarga el HTML+data.json).

**Uso local sin GitHub:** `python fetch_cps.py` (con las variables CPS_USER/CPS_PASS
en tu entorno) y abre `dashboard_cps.html` en la misma carpeta.
