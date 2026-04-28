// ============================================================
//  CLASIFICADOR METRORED-BUPA  |  Cloudflare Worker
//  OCR desde navegador -> Worker solo sube a SharePoint
// ============================================================
import { PDFDocument } from 'pdf-lib';

const LOGO_URL = 'https://raw.githubusercontent.com/MetroredEC/metrored-clasificador/main/public/logo-metrored.jpg';
const FAV_URL  = 'https://raw.githubusercontent.com/MetroredEC/metrored-clasificador/main/public/fav_icon.jpg';

function buildHTML() {
  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Clasificador BUPA | Metrored</title>
<link rel="icon" type="image/jpeg" href="${FAV_URL}">
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{--blue:#29ABE2;--blue-dark:#1a8fc0;--gray:#f4f6f9;--text:#1a2235;--sub:#6b7280}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--gray);min-height:100vh}
.topbar{background:#fff;border-bottom:2px solid #e8edf2;padding:0 32px;height:64px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.topbar-logo{height:36px}
.topbar-right{display:flex;align-items:center;gap:12px}
.btn-icon{background:none;border:1.5px solid #e5e7eb;border-radius:8px;padding:7px 14px;cursor:pointer;color:#6b7280;font-size:13px;font-weight:600;display:flex;align-items:center;gap:6px;transition:all .2s}
.btn-icon:hover{border-color:var(--blue);color:var(--blue)}
.btn-icon svg{width:15px;height:15px;stroke:currentColor;fill:none;stroke-width:2}
.main{max-width:760px;margin:0 auto;padding:32px 20px}
.page-title{font-size:22px;font-weight:700;color:var(--text);margin-bottom:4px}
.page-sub{font-size:13px;color:var(--sub);margin-bottom:28px}
.card{background:#fff;border-radius:14px;padding:24px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,.05);border:1px solid #e8edf2}
.card-title{font-size:11px;font-weight:700;color:#94a3b8;letter-spacing:.07em;text-transform:uppercase;margin-bottom:16px;display:flex;align-items:center;gap:8px}
.card-title span{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;background:var(--blue);border-radius:5px;color:#fff;font-size:11px;font-weight:800}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:16px}
label{display:block;font-size:12px;font-weight:600;color:#374151;margin-bottom:5px}
input[type=text]{width:100%;padding:9px 13px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:13px;color:var(--text);outline:none;transition:border .2s;background:#fafbfc}
input[type=text]:focus{border-color:var(--blue);background:#fff}
.drop{border:2px dashed #cbd5e1;border-radius:12px;padding:28px;text-align:center;cursor:pointer;background:#fafbfc;position:relative;transition:all .2s}
.drop.over,.drop:hover{border-color:var(--blue);background:#f0f9ff}
.drop-icon{width:44px;height:44px;background:#e0f2fe;border-radius:10px;display:flex;align-items:center;justify-content:center;margin:0 auto 10px}
.drop p{color:var(--sub);font-size:13px;margin-top:4px}.drop strong{color:var(--blue)}
#fi{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
#fname{background:#e0f2fe;color:#0369a1;font-size:12px;font-weight:600;padding:4px 10px;border-radius:6px;display:inline-block;margin-top:8px}
.btn-main{width:100%;padding:14px;background:var(--blue);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;transition:background .2s;display:flex;align-items:center;justify-content:center;gap:10px}
.btn-main:hover{background:var(--blue-dark)}
.btn-main:disabled{opacity:.5;cursor:not-allowed}
.btn-main svg{width:18px;height:18px;stroke:currentColor;fill:none;stroke-width:2}
.prog{display:none;margin-top:20px;background:#fff;border-radius:12px;padding:18px;border:1px solid #e8edf2}
.prog-bar-bg{background:#e2e8f0;border-radius:99px;height:8px;overflow:hidden;margin:10px 0}
.prog-bar{background:var(--blue);height:100%;border-radius:99px;transition:width .3s;width:0%}
.prog-msg{font-size:13px;color:var(--sub);text-align:center}
.res{display:none;margin-top:20px}
.res-sum{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:14px 18px;color:#15803d;font-weight:700;font-size:14px;margin-bottom:12px;text-align:center}
.res-item{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:11px 16px;margin-bottom:8px;font-size:13px;color:var(--text);display:flex;align-items:center;gap:10px}
.res-item::before{content:"";width:8px;height:8px;background:#22c55e;border-radius:50%;flex-shrink:0}
.res-err{background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:14px 18px;color:#dc2626;font-size:13px;margin-top:12px}
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:200;align-items:center;justify-content:center}
.modal-overlay.open{display:flex}
.modal{background:#fff;border-radius:16px;padding:32px;max-width:460px;width:calc(100% - 40px);box-shadow:0 20px 60px rgba(0,0,0,.2)}
.modal h3{font-size:17px;font-weight:700;color:var(--text);margin-bottom:4px}
.modal p{font-size:13px;color:var(--sub);margin-bottom:24px;line-height:1.6}
.modal-field{margin-bottom:16px}
.modal-actions{display:flex;gap:12px;margin-top:24px}
.btn-save{flex:1;background:var(--blue);color:#fff;border:none;padding:11px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer}
.btn-cancel{flex:1;background:#f1f5f9;color:var(--text);border:none;padding:11px;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer}
</style>
</head>
<body>

<div class="topbar">
  <img src="${LOGO_URL}" class="topbar-logo" alt="Metrored" onerror="this.style.display='none';this.nextElementSibling.style.display='block'">
  <span style="display:none;font-size:18px;font-weight:800;color:#29ABE2">Metrored</span>
  <div class="topbar-right">
    <button class="btn-icon" onclick="openSettings()">
      <svg viewBox="0 0 24 24"><path d="M12 15a3 3 0 100-6 3 3 0 000 6z"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z"/></svg>
      Configuracion
    </button>
  </div>
</div>

<div class="main">
  <div class="page-title">Clasificador de Documentos BUPA</div>
  <p class="page-sub">Sube el PDF masivo escaneado y la aplicacion organizara automaticamente cada caso en SharePoint.</p>

  <div class="card">
    <div class="card-title"><span>1</span> Informacion del Despacho</div>
    <div class="form-row">
      <div>
        <label>Numero de Despacho</label>
        <input type="text" id="despacho" placeholder="2026-04-27">
      </div>
      <div>
        <label>Carpeta Destino en SharePoint</label>
        <input type="text" id="spFolder" placeholder="Shared Documents">
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-title"><span>2</span> Archivo PDF Masivo</div>
    <div class="drop" id="drop">
      <input type="file" id="fi" accept=".pdf" onchange="pickFile(this)">
      <div class="drop-icon">
        <svg viewBox="0 0 24 24" style="width:24px;height:24px;stroke:#29ABE2;fill:none;stroke-width:1.5"><path d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
      </div>
      <p>Arrastra el PDF aqui o <strong>haz clic para seleccionar</strong></p>
      <div id="fname" style="display:none"></div>
    </div>
  </div>

  <button class="btn-main" id="btn" onclick="run()">
    <svg viewBox="0 0 24 24"><path d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
    Clasificar y Subir a SharePoint
  </button>

  <div class="prog" id="prog">
    <div class="prog-msg" id="pmsg">Iniciando...</div>
    <div class="prog-bar-bg"><div class="prog-bar" id="bar"></div></div>
  </div>

  <div class="res" id="res"></div>
</div>

<!-- MODAL CONFIGURACION -->
<div class="modal-overlay" id="settingsModal">
  <div class="modal">
    <h3>Configuracion</h3>
    <p>Define los valores predeterminados. Se guardan en tu navegador y se cargan automaticamente cada vez que abras la app.</p>
    <div class="modal-field">
      <label>Carpeta de destino en SharePoint</label>
      <input type="text" id="cfgFolder" placeholder="Shared Documents">
    </div>
    <div class="modal-field">
      <label>Prefijo de Despacho (opcional)</label>
      <input type="text" id="cfgDespacho" placeholder="ej: 2026-04">
    </div>
    <div class="modal-actions">
      <button class="btn-cancel" onclick="closeSettings()">Cancelar</button>
      <button class="btn-save" onclick="saveSettings()">Guardar</button>
    </div>
  </div>
</div>

<script>
var file = null;

window.onload = function() {
  loadSettings();
};

function loadSettings() {
  var f = localStorage.getItem('sp_folder');
  var d = localStorage.getItem('sp_despacho');
  if (f) document.getElementById('spFolder').value = f;
  if (d) document.getElementById('despacho').value = d;
}

function openSettings() {
  document.getElementById('cfgFolder').value = localStorage.getItem('sp_folder') || '';
  document.getElementById('cfgDespacho').value = localStorage.getItem('sp_despacho') || '';
  document.getElementById('settingsModal').classList.add('open');
}

function closeSettings() {
  document.getElementById('settingsModal').classList.remove('open');
}

function saveSettings() {
  localStorage.setItem('sp_folder', document.getElementById('cfgFolder').value.trim());
  localStorage.setItem('sp_despacho', document.getElementById('cfgDespacho').value.trim());
  closeSettings();
  loadSettings();
}

function pickFile(inp) {
  file = inp.files[0];
  var fn = document.getElementById('fname');
  fn.textContent = file.name;
  fn.style.display = 'inline-block';
}

var drop = document.getElementById('drop');
drop.ondragover = function(e) { e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = function() { drop.classList.remove('over'); };
drop.ondrop = function(e) {
  e.preventDefault(); drop.classList.remove('over');
  var f = e.dataTransfer.files[0];
  if (f && f.type === 'application/pdf') {
    file = f;
    var fn = document.getElementById('fname');
    fn.textContent = f.name;
    fn.style.display = 'inline-block';
  }
};

function setBar(p, m) {
  document.getElementById('bar').style.width = p + '%';
  document.getElementById('pmsg').textContent = m;
}

async function ocrChunk(pdfBytes, endpoint, key) {
  var r = await fetch(endpoint.replace(/\\/$/, '') + '/formrecognizer/documentModels/prebuilt-read:analyze?api-version=2023-07-31', {
    method: 'POST',
    headers: { 'Ocp-Apim-Subscription-Key': key, 'Content-Type': 'application/pdf' },
    body: pdfBytes
  });
  if (!r.ok) throw new Error('OCR submit: ' + await r.text());
  var opUrl = r.headers.get('Operation-Location');
  for (var i = 0; i < 24; i++) {
    await new Promise(function(res) { setTimeout(res, 4000); });
    var p = await (await fetch(opUrl, { headers: { 'Ocp-Apim-Subscription-Key': key } })).json();
    if (p.status === 'succeeded') {
      var txt = '';
      ((p.analyzeResult && p.analyzeResult.pages) || []).forEach(function(pg) {
        txt += '<<<PAG' + pg.pageNumber + '>>>\\n';
        (pg.lines || []).forEach(function(l) { txt += l.content + '\\n'; });
      });
      return txt;
    }
    if (p.status === 'failed') throw new Error('OCR fallo en este bloque');
  }
  throw new Error('OCR timeout');
}

function extractCases(ocrText, pageOffset) {
  var parts = ocrText.split(/<<<PAG\\d+>>>/);
  var pages = parts.filter(function(p) { return p.trim().length > 10; });

  function getPat(t) {
    var m = t.match(/Paciente:\\s*([A-Z\xc0-\xd6\xd8-\xde][A-Z\xc0-\xd6\xd8-\xde\\s]{4,60}?)(?:\\s{2,}|\\s*Edad:|\\s*Hc:|\\n)/m);
    return m ? m[1].trim().replace(/\\s+/g, ' ') : null;
  }
  function getPol(t) {
    var m = t.match(/\\((\\d{5,7})\\)/);
    return m ? m[1] : 'SIN_POLIZA';
  }
  function isBupa(t) { return /AUTORIZACI[O\xd3]N DE COBERTURA|N[u\xfa]mero de P[o\xf3]liza/i.test(t); }
  function isMet(t)  { return /Estado de cuenta|METRORED|Paciente:/i.test(t); }

  var cases = []; var i = 0;
  while (i < pages.length) {
    var abs = pageOffset + i + 1;
    if (isMet(pages[i])) {
      var pat = getPat(pages[i]);
      if (!pat) { i++; continue; }
      if (i + 1 < pages.length && isBupa(pages[i + 1])) {
        cases.push({ policy_number: getPol(pages[i+1]), patient_name: pat, pages: [abs, abs+1], document_types: ['estado_cuenta_metrored', 'autorizacion_cobertura_bupa'] });
        i += 2;
      } else {
        cases.push({ policy_number: 'SIN_POLIZA', patient_name: pat, pages: [abs], document_types: ['estado_cuenta_metrored'] });
        i++;
      }
    } else if (isBupa(pages[i])) {
      var am = pages[i].match(/Asegurado:\\s*([A-Z\xc0-\xd6\xd8-\xde][A-Z\xc0-\xd6\xd8-\xde\\s]{4,60}?)(?:\\s{2,}|\\s*Fecha|\\n)/m);
      cases.push({ policy_number: getPol(pages[i]), patient_name: am ? am[1].trim().replace(/\\s+/g, ' ') : 'PACIENTE_' + abs, pages: [abs], document_types: ['autorizacion_cobertura_bupa'] });
      i++;
    } else { i++; }
  }
  return cases;
}

async function run() {
  var despacho = document.getElementById('despacho').value.trim();
  var spFolder = document.getElementById('spFolder').value.trim();
  if (!despacho || !spFolder || !file) { alert('Completa todos los campos y selecciona un PDF'); return; }

  var btn = document.getElementById('btn');
  btn.disabled = true;
  document.getElementById('prog').style.display = 'block';
  document.getElementById('res').style.display = 'none';

  try {
    setBar(3, 'Obteniendo configuracion...');
    var cfg = await (await fetch('/api/config')).json();

    setBar(8, 'Leyendo PDF...');
    var buf = await file.arrayBuffer();
    var fullBytes = new Uint8Array(buf);
    var PDFDoc = PDFLib.PDFDocument;
    var fullPdf = await PDFDoc.load(fullBytes);
    var totalPages = fullPdf.getPageCount();

    var chunkSize = 4;
    var allCases = [];
    var chunks = Math.ceil(totalPages / chunkSize);

    for (var c = 0; c < chunks; c++) {
      var sp = c * chunkSize;
      var ep = Math.min(sp + chunkSize, totalPages);
      var chunkPdf = await PDFDoc.create();
      var idx = [];
      for (var p = sp; p < ep; p++) idx.push(p);
      var copied = await chunkPdf.copyPages(fullPdf, idx);
      copied.forEach(function(pg) { chunkPdf.addPage(pg); });
      var cb = await chunkPdf.save();
      setBar(10 + Math.round((c / chunks) * 55), 'OCR bloque ' + (c+1) + '/' + chunks + '...');
      var ocrText = await ocrChunk(cb, cfg.ocrEndpoint, cfg.ocrKey);
      allCases = allCases.concat(extractCases(ocrText, sp));
      // Respetar limite de rate del plan F0 (1 req / 6 seg)
      if (c < chunks - 1) {
        setBar(10 + Math.round(((c+1) / chunks) * 55), 'Esperando rate limit OCR (' + (c+1) + '/' + chunks + ')...');
        await new Promise(function(res) { setTimeout(res, 7000); });
      }
    }

    setBar(68, 'Identificados ' + allCases.length + ' casos. Subiendo a SharePoint...');

    // Send cases in batches of 5 to avoid Worker subrequest limit
    var batchSize = 5;
    var allResults = [];
    var totalBatches = Math.ceil(allCases.length / batchSize);
    var despachoCreated = false;

    for (var b = 0; b < totalBatches; b++) {
      var batchCases = allCases.slice(b * batchSize, (b + 1) * batchSize);
      var pct = 68 + Math.round((b / totalBatches) * 28);
      setBar(pct, 'Subiendo lote ' + (b+1) + '/' + totalBatches + ' a SharePoint...');

      var fd = new FormData();
      fd.append('pdf', new Blob([fullBytes], { type: 'application/pdf' }), file.name);
      fd.append('despacho', despacho);
      fd.append('spFolder', spFolder);
      fd.append('cases', JSON.stringify(batchCases));
      fd.append('despachoCreated', despachoCreated ? '1' : '0');

      var resp = await fetch('/api/upload', { method: 'POST', body: fd });
      var data = await resp.json();
      if (data.error) throw new Error('Lote ' + (b+1) + ': ' + data.error);
      despachoCreated = true;
      allResults = allResults.concat(data.cases);
    }

    setBar(100, 'Completado');
    var html = '<div class="res-sum">' + allResults.length + ' casos procesados &mdash; Despacho: ' + despacho + '</div>';
    allResults.forEach(function(c) {
      html += '<div class="res-item"><strong>' + c.folder + '</strong>&nbsp;&mdash;&nbsp;pags: ' + c.pages.join(', ') + '</div>';
    });
    var res = document.getElementById('res');
    res.innerHTML = html;
    res.style.display = 'block';

  } catch(e) {
    document.getElementById('res').innerHTML = '<div class="res-err">ERROR: ' + e.message + '</div>';
    document.getElementById('res').style.display = 'block';
    setBar(0, 'Error');
  }
  btn.disabled = false;
}
</script>
</body>
</html>`;
}

export default {
  async fetch(request, env) {
    const url  = new URL(request.url);
    const cors = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' };
    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });
    if (url.pathname === '/') return new Response(buildHTML(), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });

    if (url.pathname === '/api/config') {
      return new Response(JSON.stringify({ ocrEndpoint: env.DOCAI_ENDPOINT, ocrKey: env.DOCAI_KEY }), { headers: { ...cors, 'Content-Type': 'application/json' } });
    }

    if (url.pathname === '/api/upload' && request.method === 'POST') {
      try {
        const result = await handleUpload(request, env);
        return new Response(JSON.stringify(result), { headers: { ...cors, 'Content-Type': 'application/json' } });
      } catch (err) {
        return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { ...cors, 'Content-Type': 'application/json' } });
      }
    }

    return new Response('Not found', { status: 404 });
  }
};

async function handleUpload(request, env) {
  const form      = await request.formData();
  const pdfFile   = form.get('pdf');
  const despacho  = form.get('despacho');
  const spFolder  = form.get('spFolder');
  const casesJson = form.get('cases');

  const allCases = JSON.parse(casesJson);
  if (!allCases || allCases.length === 0) throw new Error('No se recibieron casos para procesar');

  const pdfBytes = new Uint8Array(await pdfFile.arrayBuffer());
  const pdfDoc   = await PDFDocument.load(pdfBytes);
  const total    = pdfDoc.getPageCount();

  const spToken  = await getSharePointToken(env.AZURE_TENANT_ID, env.AZURE_CLIENT_ID, env.AZURE_CLIENT_SECRET);
  const siteInfo = await getSharePointSite(spToken, env.SP_HOSTNAME, env.SP_SITE_PATH);
  const despachoPath = spFolder + '/DESPACHO-' + despacho;
  const despachoCreated = form.get('despachoCreated') === '1';
  if (!despachoCreated) {
    await createSharePointFolder(spToken, siteInfo.driveId, despachoPath);
  }

  const results = [];
  for (const caso of allCases) {
    const patient  = sanitize(caso.patient_name || 'DESCONOCIDO');
    const policy   = (caso.policy_number || 'SIN_POLIZA').toString().trim();
    const folder   = policy + ' - ' + patient;
    const casePath = despachoPath + '/' + folder;
    const pages    = caso.pages || [];
    const types    = caso.document_types || [];

    await createSharePointFolder(spToken, siteInfo.driveId, casePath);

    // Build a single PDF with all pages of this case
    const casePdf = await PDFDocument.create();
    for (let i = 0; i < pages.length; i++) {
      const idx = pages[i] - 1;
      if (idx < 0 || idx >= total) continue;
      const [pg] = await casePdf.copyPages(pdfDoc, [idx]);
      casePdf.addPage(pg);
    }
    const caseBytes = await casePdf.save();
    const caseName  = policy + '_' + patient + '.pdf';
    await uploadSharePointFile(spToken, siteInfo.driveId, casePath + '/' + caseName, caseBytes);
    results.push({ folder, pages, status: 'OK' });
  }

  return { success: true, despacho, cases_processed: results.length, cases: results };
}

async function getSharePointToken(tenantId, clientId, clientSecret) {
  const body = 'grant_type=client_credentials&client_id=' + encodeURIComponent(clientId) + '&client_secret=' + encodeURIComponent(clientSecret) + '&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default';
  const res  = await fetch('https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  const tok  = await res.json();
  if (!tok.access_token) throw new Error('Azure auth: ' + JSON.stringify(tok));
  return tok.access_token;
}

async function getSharePointSite(token, hostname, sitePath) {
  const sr = await fetch('https://graph.microsoft.com/v1.0/sites/' + hostname + ':/' + sitePath, { headers: { 'Authorization': 'Bearer ' + token } });
  const s  = await sr.json();
  if (!s.id) throw new Error('Sitio SharePoint no encontrado: ' + JSON.stringify(s));
  const dr = await fetch('https://graph.microsoft.com/v1.0/sites/' + s.id + '/drive', { headers: { 'Authorization': 'Bearer ' + token } });
  const d  = await dr.json();
  if (!d.id) throw new Error('Drive no encontrado: ' + JSON.stringify(d));
  return { siteId: s.id, driveId: d.id };
}

async function createSharePointFolder(token, driveId, path) {
  const parts = path.split('/').filter(Boolean);
  let pid = 'root';
  for (const part of parts) {
    const r = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + pid + '/children', {
      method: 'POST', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: part, folder: {} })
    });
    const d = await r.json();
    if (d.error && d.error.code === 'nameAlreadyExists') {
      const er = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + pid + ':/' + encodeURIComponent(part), { headers: { 'Authorization': 'Bearer ' + token } });
      const ed = await er.json(); pid = ed.id;
    } else if (!d.id) { throw new Error('Error carpeta "' + part + '": ' + JSON.stringify(d)); }
    else { pid = d.id; }
  }
  return pid;
}

async function uploadSharePointFile(token, driveId, filePath, bytes) {
  const parts  = filePath.split('/').filter(Boolean);
  const name   = parts.pop();
  const folder = parts.join('/');
  const fr = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/root:/' + folder, { headers: { 'Authorization': 'Bearer ' + token } });
  const fd = await fr.json();
  if (!fd.id) throw new Error('Carpeta upload no encontrada: ' + folder);
  const ur = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + fd.id + ':/' + encodeURIComponent(name) + ':/content', {
    method: 'PUT', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/pdf' }, body: bytes
  });
  const ud = await ur.json();
  if (!ud.id) throw new Error('Error subiendo "' + name + '": ' + JSON.stringify(ud));
  return ud.id;
}

function sanitize(str) { return str.replace(/[<>:"/\\|?*\x00-\x1F#%{}^~\[\]`']/g, '').replace(/\s+/g, ' ').trim().substring(0, 60); }