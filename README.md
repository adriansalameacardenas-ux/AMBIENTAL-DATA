<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Extractor CPS → Excel (Automático)</title>
<script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: #0d1117; color: #e6edf3; min-height: 100vh; }
  header { background: linear-gradient(135deg, #1a2f1a, #0d3b2e); padding: 28px 20px; text-align: center; border-bottom: 2px solid #2ea043; }
  header h1 { font-size: 1.5rem; } header p { color: #9fe6b8; margin-top: 6px; font-size: .9rem; }
  main { max-width: 980px; margin: 24px auto; padding: 0 16px; display: grid; gap: 20px; }
  .card { background: #161b22; border: 1px solid #30363d; border-radius: 12px; padding: 20px; }
  .card h2 { font-size: 1.05rem; color: #58d68d; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }
  .num { background: #2ea043; color: #fff; width: 24px; height: 24px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: .8rem; font-weight: bold; }
  label { display: block; font-size: .8rem; color: #8b949e; margin: 10px 0 4px; }
  input, textarea { width: 100%; background: #0d1117; border: 1px solid #30363d; border-radius: 8px; color: #e6edf3; padding: 9px 12px; font-size: .88rem; font-family: inherit; }
  input:focus, textarea:focus { outline: none; border-color: #2ea043; }
  textarea { min-height: 100px; resize: vertical; font-family: Consolas, monospace; font-size: .8rem; }
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  .grid3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }
  button { cursor: pointer; border: none; border-radius: 8px; padding: 12px 20px; font-size: .95rem; font-weight: 600; transition: .15s; }
  .btn-auto { background: linear-gradient(135deg, #2ea043, #1f8f36); color: #fff; width: 100%; margin-top: 14px; font-size: 1.05rem; padding: 14px; }
  .btn-auto:hover { filter: brightness(1.15); }
  .btn-auto:disabled { opacity: .5; cursor: not-allowed; }
  .btn-blue { background: #1f6feb; color: #fff; width: 100%; } .btn-blue:hover { background: #388bfd; }
  .btn-gray { background: #21262d; color: #c9d1d9; border: 1px solid #30363d; } .btn-gray:hover { background: #30363d; }
  .btn-orange { background: #b45309; color: #fff; } .btn-orange:hover { background: #d97706; }
  #log { background: #0d1117; border: 1px solid #30363d; border-radius: 8px; padding: 12px; font-family: Consolas, monospace; font-size: .78rem; max-height: 200px; overflow-y: auto; white-space: pre-wrap; }
  .ok { color: #3fb950; } .err { color: #f85149; } .warn { color: #d29922; } .info { color: #58a6ff; }
  table { width: 100%; border-collapse: collapse; font-size: .78rem; margin-top: 10px; }
  th, td { border: 1px solid #30363d; padding: 5px 8px; text-align: left; white-space: nowrap; }
  th { background: #1c2128; color: #58d68d; position: sticky; top: 0; }
  .preview-wrap { max-height: 260px; overflow: auto; border-radius: 8px; }
  .pill { display: inline-block; border-radius: 20px; padding: 2px 12px; font-size: .75rem; margin: 2px 4px 2px 0; cursor: pointer; }
  .pill-paste { background: #1f6feb22; color: #58a6ff; border: 1px solid #1f6feb55; }
  .sheet-tag { background: #2ea04322; color: #3fb950; border: 1px solid #2ea04355; cursor: default; }
  footer { text-align: center; color: #484f58; font-size: .75rem; padding: 20px; }
  .hint { font-size: .78rem; color: #8b949e; margin-top: 8px; line-height: 1.5; }
  .spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid #fff3; border-top-color: #fff; border-radius: 50%; animation: spin .7s linear infinite; vertical-align: -3px; margin-right: 8px; }
  @keyframes spin { to { transform: rotate(360deg); } }
  details summary { cursor: pointer; color: #58a6ff; font-size: .85rem; margin-top: 10px; }
  .cors-warn { background: #3b2d0b; border: 1px solid #d29922; border-radius: 8px; padding: 12px; font-size: .82rem; color: #e3b341; margin-top: 10px; display: none; line-height: 1.5; }
</style>
<base target="_blank">
</head>
<body>
<header>
  <h1>☀️ Extractor Chint Power → Excel <small style="font-size:.7em; color:#3fb950">v2 · automático</small></h1>
  <p>monitor.chintpowersystems.com · un botón y te baja el .xlsx</p>
</header>
<main>

  <!-- MODO AUTOMÁTICO -->
  <div class="card">
    <h2>⚡ Modo automático</h2>
    <div class="grid3">
      <div><label>URL base</label><input id="baseUrl" value="https://monitor.chintpowersystems.com"></div>
      <div><label>Usuario</label><input id="user" placeholder="tu usuario" autocomplete="username"></div>
      <div><label>Clave</label><input id="pass" type="password" placeholder="tu clave" autocomplete="current-password"></div>
    </div>
    <details>
      <summary>Opciones avanzadas (endpoints y proxy CORS)</summary>
      <div class="grid3" style="margin-top:8px">
        <div><label>Endpoint login (vacío = autodetectar)</label><input id="epLogin" placeholder="autodetección"></div>
        <div><label>Endpoint plantas (vacío = autodetectar)</label><input id="epPlants" placeholder="autodetección"></div>
        <div><label>Endpoint dispositivos (vacío = autodetectar)</label><input id="epDevices" placeholder="autodetección"></div>
      </div>
      <label>Proxy CORS (opcional, solo si el modo automático falla por CORS — ⚠️ tu clave pasaría por un tercero, úsalo bajo tu responsabilidad)</label>
      <input id="corsProxy" placeholder="https://api.allorigins.win/raw?url=  (déjalo vacío para no usar)">
    </details>
    <button class="btn-auto" id="btnAuto" onclick="modoAutomatico()">🚀 CONECTAR, EXTRAER Y DESCARGAR EXCEL</button>
    <div class="cors-warn" id="corsWarn">
      ⚠️ <b>El portal bloqueó la conexión automática (CORS).</b> Es lo habitual en este tipo de plataformas.
      No hay problema: usa el <b>Plan B</b> de abajo (pegar el JSON desde DevTools — toma 1 minuto) y el Excel se descarga igual.
      <br><br>
      <b>Cómo:</b> 1) Entra al portal e inicia sesión · 2) F12 → Network → Fetch/XHR ·
      3) click en las peticiones (plant/device/energy) → Copy response · 4) pega abajo → Procesar → Descargar.
    </div>
  </div>

  <!-- PLAN B -->
  <div class="card">
    <h2><span class="num">B</span> Plan B: pegar JSON (funciona siempre)</h2>
    <label>📋 PLANTAS <span class="pill pill-paste" onclick="pegar('taPlants')">📥 pegar</span></label>
    <textarea id="taPlants" placeholder='{"data":{"list":[...]}}'></textarea>
    <label>📋 DISPOSITIVOS <span class="pill pill-paste" onclick="pegar('taDevices')">📥 pegar</span></label>
    <textarea id="taDevices" placeholder='{"data":{"list":[...]}}'></textarea>
    <label>📋 HISTÓRICO / RESUMEN (varios JSON separados por línea en blanco) <span class="pill pill-paste" onclick="pegar('taHist')">📥 pegar</span></label>
    <textarea id="taHist" placeholder='Uno o varios JSON'></textarea>
    <button class="btn-blue" style="margin-top:12px" onclick="procesarPegado(true)">⚙️ Procesar y descargar Excel</button>
  </div>

  <!-- LOG -->
  <div class="card">
    <h2>📟 Registro</h2>
    <div id="log"><span class="info">Pulsa el botón verde para empezar, o pega JSON en el Plan B.</span></div>
  </div>

  <!-- PREVIEW -->
  <div class="card" id="cardPreview" style="display:none">
    <h2>👀 Vista previa</h2>
    <div id="chips"></div>
    <div id="preview" class="preview-wrap"></div>
  </div>

  <!-- DESCARGA -->
  <div class="card" id="cardDownload" style="display:none">
    <h2>📥 Descargar</h2>
    <div class="grid2">
      <div>
        <label>Nombre del archivo</label>
        <input id="fileName" value="chint_power_data">
      </div>
      <div style="display:flex; align-items:flex-end; gap:10px">
        <button class="btn-gray" onclick="descargarExcel(false)">⬇️ .xlsx</button>
        <button class="btn-orange" onclick="descargarExcel(true)">🗜️ .zip</button>
      </div>
    </div>
  </div>

</main>
<footer>Herramienta local: ningún dato sale de tu navegador · SheetJS + JSZip</footer>

<script>
/* ============================ ESTADO ============================ */
const SHEETS = {};
let TOKEN = null;

const LOGIN_PATHS = ['/api/login','/api/user/login','/api/v1/login','/api/v1/user/login',
  '/api/auth/login','/api/v1/auth/login','/login','/api/user/signin'];
const PLANT_PATHS = ['/api/plant/list','/api/plant/getPlantList','/api/v1/plant/list',
  '/api/v1/plants','/api/plant','/api/v1/plant/page','/api/station/list','/api/v1/station/list'];
const DEVICE_PATHS = ['/api/device/list','/api/v1/device/list','/api/device','/api/v1/device/page'];
const HIST_PATHS = ['/api/device/history','/api/v1/device/history','/api/plant/energy',
  '/api/v1/plant/energy','/api/v1/energy','/api/plant/chart','/api/v1/plant/chart'];

/* ============================ UTILIDADES ============================ */
function log(msg, cls='info') {
  const el = document.getElementById('log');
  el.innerHTML += `<div class="${cls}">[${new Date().toLocaleTimeString()}] ${msg}</div>`;
  el.scrollTop = el.scrollHeight;
}
function aplanar(obj, prefix='', out={}) {
  if (obj === null || obj === undefined) return out;
  if (typeof obj !== 'object') { out[prefix] = obj; return out; }
  if (Array.isArray(obj)) { out[prefix] = JSON.stringify(obj); return out; }
  for (const k of Object.keys(obj)) aplanar(obj[k], prefix ? `${prefix}.${k}` : k, out);
  return out;
}
function extraerLista(data) {
  if (Array.isArray(data)) return data;
  if (data && typeof data === 'object') {
    for (const k of ['list','rows','records','data','result','items','plants','devices']) {
      if (Array.isArray(data[k])) return data[k];
      if (data[k] && typeof data[k] === 'object') {
        const r = extraerLista(data[k]);
        if (r) return r;
      }
    }
  }
  return null;
}
function aFilas(data) {
  const lista = extraerLista(data);
  if (lista) return lista.filter(x => x && typeof x === 'object').map(x => aplanar(x));
  if (data && typeof data === 'object') return [aplanar(data)];
  return [];
}
function columnas(filas) {
  const s = new Set(); filas.forEach(f => Object.keys(f).forEach(k => s.add(k)));
  return [...s];
}
function agregarHoja(nombre, data) {
  const filas = aFilas(data);
  if (!filas.length) { log(`⚠️ "${nombre}": sin filas`, 'warn'); return false; }
  let n = nombre, i = 2;
  while (SHEETS[n]) n = `${nombre}_${i++}`;
  SHEETS[n] = filas;
  log(`✅ Hoja "${n}": ${filas.length} filas × ${columnas(filas).length} columnas`, 'ok');
  renderPreview();
  return true;
}
async function pegar(id) {
  try {
    document.getElementById(id).value = await navigator.clipboard.readText();
    log('📥 Pegado desde el portapapeles');
  } catch(e) { log('No se pudo leer el portapapeles: ' + e.message, 'err'); }
}
function esCORS(e) { return /failed to fetch|networkerror|load failed|cors/i.test(String(e)); }

/* ============================ API ============================ */
function construirUrl(path) {
  let base = document.getElementById('baseUrl').value.replace(/\/+$/, '');
  let url = path.startsWith('http') ? path : base + path;
  const proxy = document.getElementById('corsProxy').value.trim();
  if (proxy) url = proxy + encodeURIComponent(url);
  return url;
}
async function api(path, opts={}) {
  const headers = { 'Accept': 'application/json, text/plain, */*' };
  if (opts.body) headers['Content-Type'] = 'application/json';
  if (TOKEN) headers['Authorization'] = 'Bearer ' + TOKEN;
  const r = await fetch(construirUrl(path), { ...opts, headers });
  if (!r.ok) throw new Error(`HTTP ${r.statusCode || r.status} en ${path}`);
  return r.json();
}
function buscarToken(d) {
  if (!d || typeof d !== 'object') return null;
  for (const k of ['token','access_token','accessToken','jwt']) if (d[k]) return d[k];
  for (const k of ['data','result','body']) { const t = buscarToken(d[k]); if (t) return t; }
  return null;
}
async function intentarLogin() {
  const u = document.getElementById('user').value.trim();
  const p = document.getElementById('pass').value;
  if (!u || !p) throw new Error('Completa usuario y clave.');
  const manual = document.getElementById('epLogin').value.trim();
  const paths = manual ? [manual] : LOGIN_PATHS;
  const formas = [{username:u,password:p},{userName:u,password:p},{account:u,password:p},{email:u,password:p}];
  let huboCORS = false, ultErr = null;
  for (const path of paths) {
    for (const body of formas) {
      try {
        const data = await api(path, { method:'POST', body: JSON.stringify(body) });
        const t = buscarToken(data);
        if (t) { TOKEN = t; log('🔑 Login OK en ' + path + ' (token)', 'ok'); return true; }
        log('Login en ' + path + ' respondió; se continúa con la sesión (cookies).', 'ok');
        return true;
      } catch(e) {
        if (esCORS(e)) { huboCORS = true; break; }
        ultErr = e;
      }
    }
    if (huboCORS) break;
  }
  if (huboCORS) throw new Error('CORS');
  throw ultErr || new Error('Login falló: revisa credenciales o el endpoint con F12.');
}
async function autodetectar(paths, etiqueta) {
  const manual = (etiqueta === 'plantas' && document.getElementById('epPlants').value.trim())
              || (etiqueta === 'dispositivos' && document.getElementById('epDevices').value.trim());
  const lista = manual ? [manual] : paths;
  for (const path of lista) {
    try {
      const data = await api(path, {});
      const filas = aFilas(data);
      if (filas.length) { log('📡 Endpoint de ' + etiqueta + ' detectado: ' + path, 'ok'); return { path, data, filas }; }
    } catch(e) { if (!esCORS(e)) log('  probando ' + path + '… ' + e.message, 'info'); }
  }
  return null;
}

/* ============================ MODO AUTOMÁTICO ============================ */
async function modoAutomatico() {
  const btn = document.getElementById('btnAuto');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Conectando y extrayendo…';
  document.getElementById('corsWarn').style.display = 'none';
  try {
    log('🚀 Inicio del modo automático…');
    await intentarLogin();

    const pl = await autodetectar(PLANT_PATHS, 'plantas');
    if (pl) agregarHoja('Plantas', pl.data);

    const dv = await autodetectar(DEVICE_PATHS, 'dispositivos');
    if (dv) agregarHoja('Dispositivos', dv.data);

    // histórico por dispositivo (día/mes/año), solo si hay endpoint y SNs
    const hp = await autodetectar(HIST_PATHS, 'histórico');
    if (hp && dv) {
      const snCol = columnas(dv.filas).find(c => /sn|serial/i.test(c));
      if (snCol) {
        let n = 0;
        for (const f of dv.filas.slice(0, 10)) {
          const sn = f[snCol];
          if (!sn) continue;
          for (const tipo of ['day','month','year']) {
            try {
              const h = await api(hp.path, {});
              if (agregarHoja(`Hist_${String(sn).slice(0,10)}_${tipo}`, h)) n++;
            } catch(e) { /* sin datos de ese tipo */ }
            break; // muchos portales devuelven todo en una llamada; probamos 1 vez por SN
          }
        }
        if (!n) log('Histórico: el endpoint responde pero sin series por SN; agrégalo manual en Plan B si lo necesitas.', 'warn');
      }
    }

    if (Object.keys(SHEETS).length) {
      log('🎉 Datos listos. Descargando Excel automáticamente…', 'ok');
      descargarExcel(false);
    } else {
      log('⚠️ No se obtuvieron datos. Intenta el Plan B (pegar JSON).', 'warn');
    }
  } catch(e) {
    if (String(e.message) === 'CORS' || esCORS(e)) {
      log('❌ Conexión bloqueada por CORS (el navegador no puede hablar con el portal desde esta página).', 'err');
      document.getElementById('corsWarn').style.display = 'block';
      document.getElementById('corsWarn').scrollIntoView({ behavior:'smooth' });
    } else {
      log('❌ ' + e.message, 'err');
    }
  } finally {
    btn.disabled = false;
    btn.innerHTML = '🚀 CONECTAR, EXTRAER Y DESCARGAR EXCEL';
  }
}

/* ============================ PLAN B ============================ */
function procesarPegado(autoDownload) {
  const p = document.getElementById('taPlants').value.trim();
  const d = document.getElementById('taDevices').value.trim();
  const h = document.getElementById('taHist').value.trim();
  if (p) { try { agregarHoja('Plantas', JSON.parse(p)); } catch(e){ log('❌ JSON plantas inválido: '+e.message,'err'); } }
  if (d) { try { agregarHoja('Dispositivos', JSON.parse(d)); } catch(e){ log('❌ JSON dispositivos inválido: '+e.message,'err'); } }
  if (h) {
    h.split(/\n\s*\n/).map(b => b.trim()).filter(Boolean).forEach((b, i, arr) => {
      try { agregarHoja(arr.length > 1 ? `Historico_${i+1}` : 'Historico', JSON.parse(b)); }
      catch(e){ log(`❌ Bloque histórico ${i+1} inválido: ${e.message}`,'err'); }
    });
  }
  if (autoDownload && Object.keys(SHEETS).length) descargarExcel(false);
  if (!Object.keys(SHEETS).length) log('Nada que procesar aún.', 'warn');
}

/* ============================ PREVIEW ============================ */
function renderPreview() {
  const nombres = Object.keys(SHEETS);
  document.getElementById('cardPreview').style.display = nombres.length ? 'block' : 'none';
  document.getElementById('cardDownload').style.display = nombres.length ? 'block' : 'none';
  document.getElementById('chips').innerHTML = nombres.map(n =>
    `<span class="pill sheet-tag">${n} · ${SHEETS[n].length} filas</span>`).join('');
  const cont = document.getElementById('preview');
  cont.innerHTML = '';
  nombres.forEach(n => {
    const filas = SHEETS[n];
    const cols = columnas(filas).slice(0, 10);
    let h = `<h3 style="margin:14px 0 4px; color:#c9d1d9; font-size:.9rem">${n}</h3><table><thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}</tr></thead><tbody>`;
    filas.slice(0, 8).forEach(f => { h += '<tr>' + cols.map(c => `<td>${String(f[c] ?? '').slice(0,40)}</td>`).join('') + '</tr>'; });
    cont.innerHTML += h + '</tbody></table>';
  });
}

/* ============================ EXCEL / ZIP ============================ */
function descargarExcel(zip) {
  const nombres = Object.keys(SHEETS);
  if (!nombres.length) { log('No hay datos para exportar.', 'warn'); return; }
  const wb = XLSX.utils.book_new();
  nombres.forEach(n => {
    const cols = columnas(SHEETS[n]);
    const aoa = [cols, ...SHEETS[n].map(f => cols.map(c => f[c] ?? ''))];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = cols.map(c => ({ wch: Math.min(Math.max(c.length + 2, 12), 40) }));
    XLSX.utils.book_append_sheet(wb, ws, n.replace(/[^\w]/g,'_').slice(0,31));
  });
  const base = (document.getElementById('fileName').value || 'chint_power_data').replace(/[^\w\-]/g,'_');
  if (!zip) {
    XLSX.writeFile(wb, base + '.xlsx');
    log('⬇️ Excel descargado: ' + base + '.xlsx', 'ok');
  } else {
    const wbout = XLSX.write(wb, { bookType:'xlsx', type:'array' });
    const z = new JSZip();
    z.file(base + '.xlsx', wbout);
    z.generateAsync({ type:'blob' }).then(b => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(b);
      a.download = base + '.zip';
      a.click();
      log('🗜️ ZIP descargado: ' + base + '.zip', 'ok');
    });
  }
}
</script>
</body>
</html>
