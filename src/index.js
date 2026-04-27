// ============================================================
//  CLASIFICADOR METRORED-BUPA  |  Cloudflare Worker
//  Extrae datos con IA, crea carpetas y sube PDFs a Drive
// ============================================================

import { PDFDocument } from 'pdf-lib';

// ─────────────────────────── HTML UI ───────────────────────────
const HTML_PAGE = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Clasificador BUPA-Metrored</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:linear-gradient(135deg,#1a56db 0%,#7e3af2 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:#fff;border-radius:18px;padding:40px;max-width:620px;width:100%;box-shadow:0 25px 70px rgba(0,0,0,.25)}
h1{font-size:22px;color:#111827;text-align:center;margin-bottom:6px}
.subtitle{color:#6b7280;font-size:13px;text-align:center;margin-bottom:30px}
.badge{display:inline-block;background:#ede9fe;color:#6d28d9;font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;margin-left:6px;vertical-align:middle}
.section{background:#f9fafb;border-radius:10px;padding:20px;margin-bottom:18px}
.section-title{font-size:12px;font-weight:700;color:#6b7280;letter-spacing:.06em;text-transform:uppercase;margin-bottom:14px}
label{display:block;font-size:13px;font-weight:600;color:#374151;margin-bottom:6px}
label .hint{font-weight:400;color:#9ca3af;font-size:11px;margin-left:4px}
input[type=text]{width:100%;padding:9px 13px;border:1.5px solid #e5e7eb;border-radius:8px;font-size:13px;outline:none;transition:border-color .2s}
input[type=text]:focus{border-color:#7c3aed}
.row{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.file-drop{border:2px dashed #d1d5db;border-radius:10px;padding:26px;text-align:center;cursor:pointer;transition:all .2s;background:#fff;position:relative}
.file-drop:hover,.file-drop.over{border-color:#7c3aed;background:#f5f3ff}
.file-drop p{color:#6b7280;font-size:13px;margin-top:8px}
.file-drop strong{color:#7c3aed}
.fname{color:#111827;font-weight:600;font-size:13px;margin-top:6px;display:none}
#fi{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.btn{width:100%;padding:13px;background:linear-gradient(135deg,#1a56db,#7e3af2);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;margin-top:4px;transition:opacity .2s;letter-spacing:.02em}
.btn:disabled{opacity:.55;cursor:not-allowed}
.btn:not(:disabled):hover{opacity:.9}
.progress{display:none;margin-top:22px;background:#f3f4f6;border-radius:10px;padding:18px}
.bar-bg{background:#e5e7eb;border-radius:99px;height:7px;overflow:hidden;margin:10px 0}
.bar{background:linear-gradient(90deg,#1a56db,#7e3af2);height:100%;border-radius:99px;transition:width .4s;width:0%}
.pmsg{font-size:13px;color:#6b7280;text-align:center}
.results{display:none;margin-top:20px}
.ok-box{background:#ecfdf5;border-left:4px solid #10b981;border-radius:6px;padding:12px 16px;margin-bottom:8px;display:flex;align-items:center;gap:10px;font-size:13px;color:#065f46}
.err-box{background:#fef2f2;border-left:4px solid #ef4444;border-radius:6px;padding:12px 16px;color:#991b1b;font-size:13px;margin-top:12px}
.summary{text-align:center;background:#ede9fe;border-radius:8px;padding:12px;margin-bottom:14px;color:#5b21b6;font-weight:700;font-size:14px}
</style>
</head>
<body>
<div class="card">
  <div style="text-align:center;margin-bottom:20px">
    <svg width="46" height="46" viewBox="0 0 46 46" fill="none" xmlns="http://www.w3.org/2000/svg" style="display:inline-block">
      <rect width="46" height="46" rx="12" fill="url(#g1)"/>
      <defs><linearGradient id="g1" x1="0" y1="0" x2="46" y2="46"><stop offset="0%" stop-color="#1a56db"/><stop offset="100%" stop-color="#7e3af2"/></linearGradient></defs>
      <path d="M23 11v24M11 23h24" stroke="white" stroke-width="3.5" stroke-linecap="round"/>
    </svg>
    <h1>Clasificador BUPA-Metrored <span class="badge">IA</span></h1>
    <p class="subtitle">Organiza escaneados masivos en Google Drive automáticamente</p>
  </div>

  <div class="section">
    <div class="section-title">📋 Información del Lote</div>
    <div class="row">
      <div>
        <label>Número de Despacho <span class="hint">p. ej: 2026-04-08</span></label>
        <input type="text" id="despacho" placeholder="2026-04-08">
      </div>
      <div>
        <label>Carpeta Drive Destino <span class="hint">URL o ID</span></label>
        <input type="text" id="driveUrl" placeholder="https://drive.google.com/drive/folders/...">
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">📄 PDF Masivo</div>
    <div class="file-drop" id="drop">
      <input type="file" id="fi" accept=".pdf" onchange="pickFile(this)">
      <svg width="36" height="36" fill="none" stroke="#9ca3af" stroke-width="1.5" viewBox="0 0 24 24">
        <path d="M9 12h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
      </svg>
      <p>Arrastra el PDF aquí o <strong>haz clic para seleccionar</strong></p>
      <div class="fname" id="fname"></div>
    </div>
  </div>

  <button class="btn" id="btn" onclick="run()">🚀 Clasificar y Subir a Drive</button>

  <div class="progress" id="prog">
    <div class="pmsg" id="pmsg">Iniciando...</div>
    <div class="bar-bg"><div class="bar" id="bar"></div></div>
  </div>

  <div class="results" id="res"></div>
</div>

<script>
let file = null;

function pickFile(inp){
  file = inp.files[0];
  document.getElementById('fname').textContent = file.name;
  document.getElementById('fname').style.display = 'block';
}

const drop = document.getElementById('drop');
drop.ondragover = e=>{ e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = ()=> drop.classList.remove('over');
drop.ondrop = e=>{
  e.preventDefault(); drop.classList.remove('over');
  const f = e.dataTransfer.files[0];
  if(f && f.type==='application/pdf'){
    file = f;
    document.getElementById('fname').textContent = f.name;
    document.getElementById('fname').style.display = 'block';
  }
};

function setBar(pct, msg){
  document.getElementById('bar').style.width = pct+'%';
  document.getElementById('pmsg').textContent = msg;
}

async function run(){
  const despacho = document.getElementById('despacho').value.trim();
  const driveRaw = document.getElementById('driveUrl').value.trim();
  if(!despacho) return alert('Ingresa el número de despacho');
  if(!driveRaw)  return alert('Ingresa la URL o ID de la carpeta de Drive');
  if(!file)      return alert('Selecciona un archivo PDF');

  const m = driveRaw.match(/\\/folders\\/([a-zA-Z0-9_\\-]+)/);
  const folderId = m ? m[1] : driveRaw;

  const btn = document.getElementById('btn');
  btn.disabled = true;
  document.getElementById('prog').style.display = 'block';
  document.getElementById('res').style.display = 'none';
  setBar(8, '📄 Leyendo PDF...');

  try {
    const fd = new FormData();
    fd.append('pdf', file);
    fd.append('despacho', despacho);
    fd.append('driveFolderId', folderId);

    setBar(20, '🤖 Extrayendo datos con IA (puede tardar 30-60 seg)...');

    const resp = await fetch('/api/process', { method:'POST', body:fd });

    setBar(75, '📁 Creando carpetas y subiendo archivos a Drive...');

    const data = await resp.json();
    if(!resp.ok || data.error) throw new Error(data.error || 'Error desconocido');

    setBar(100, '✅ Proceso completado');
    showResults(data);
  } catch(e){
    document.getElementById('res').innerHTML = '<div class="err-box">❌ '+e.message+'</div>';
    document.getElementById('res').style.display = 'block';
    setBar(0, 'Error al procesar');
  }
  btn.disabled = false;
}

function showResults(data){
  let html = '<div class="summary">✅ '+data.cases_processed+' casos procesados — Despacho: '+data.despacho+'</div>';
  data.cases.forEach(c=>{
    html += '<div class="ok-box"><svg width="16" height="16" fill="#10b981" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/></svg><div><strong>'+c.folder+'</strong> <span style="opacity:.7">(págs. '+c.pages.join(', ')+')</span></div></div>';
  });
  const res = document.getElementById('res');
  res.innerHTML = html;
  res.style.display = 'block';
}
</script>
</body>
</html>`;

// ─────────────────────────── WORKER ENTRY ───────────────────────────
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });

    // ── Serve UI
    if (url.pathname === '/') {
      return new Response(HTML_PAGE, { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }

    // ── Main processing endpoint
    if (url.pathname === '/api/process' && request.method === 'POST') {
      try {
        const result = await handleProcess(request, env);
        return new Response(JSON.stringify(result), {
          headers: { ...cors, 'Content-Type': 'application/json' }
        });
      } catch (err) {
        console.error('Process error:', err);
        return new Response(JSON.stringify({ error: err.message }), {
          status: 500,
          headers: { ...cors, 'Content-Type': 'application/json' }
        });
      }
    }

    return new Response('Not found', { status: 404 });
  }
};

// ─────────────────────────── MAIN HANDLER ───────────────────────────
async function handleProcess(request, env) {
  // 1. Parse form
  const form = await request.formData();
  const pdfFile     = form.get('pdf');
  const despachoNum = form.get('despacho');
  const rootFolderId = form.get('driveFolderId');

  if (!pdfFile || !despachoNum || !rootFolderId) {
    throw new Error('Faltan campos: pdf, despacho, driveFolderId');
  }

  const pdfBytes  = new Uint8Array(await pdfFile.arrayBuffer());
  const pdfBase64 = uint8ToBase64(pdfBytes);

  // 2. Extract cases with Claude AI
  const cases = await extractCasesWithClaude(pdfBase64, env.ANTHROPIC_API_KEY);
  if (!Array.isArray(cases) || cases.length === 0) {
    throw new Error('No se pudieron extraer casos del PDF. Verifica que el archivo sea correcto.');
  }

  // 3. Google Drive auth
  const serviceAccount = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT);
  const driveToken = await getGoogleDriveToken(serviceAccount);

  // 4. Create parent DESPACHO folder
  const despachoFolderId = await createDriveFolder(
    `DESPACHO-${despachoNum}`,
    rootFolderId,
    driveToken
  );

  // 5. Load PDF for splitting
  const pdfDoc = await PDFDocument.load(pdfBytes);
  const results = [];

  // 6. Process each case
  for (const caso of cases) {
    const patientClean  = sanitize(caso.patient_name || 'PACIENTE_DESCONOCIDO');
    const policyNum     = (caso.policy_number || 'SIN_POLIZA').toString().trim();
    const folderName    = `${policyNum} - ${patientClean}`;

    // Create subfolder
    const caseFolderId = await createDriveFolder(folderName, despachoFolderId, driveToken);

    // Upload each page as individual PDF
    const pages = caso.pages || [];
    const docTypes = caso.document_types || [];

    for (let i = 0; i < pages.length; i++) {
      const pageIndex = pages[i] - 1; // 0-based
      if (pageIndex < 0 || pageIndex >= pdfDoc.getPageCount()) continue;

      const singlePdf = await PDFDocument.create();
      const [copied]  = await singlePdf.copyPages(pdfDoc, [pageIndex]);
      singlePdf.addPage(copied);
      const singleBytes = await singlePdf.save();

      const docType  = docTypes[i] || (i === 0 ? 'estado_cuenta_metrored' : 'autorizacion_cobertura_bupa');
      const fileName = `${i + 1}_${docType}.pdf`;

      await uploadToDrive(fileName, singleBytes, caseFolderId, driveToken);
    }

    results.push({ folder: folderName, pages, status: 'OK' });
  }

  return {
    success: true,
    despacho: despachoNum,
    cases_processed: results.length,
    cases: results
  };
}

// ─────────────────────────── CLAUDE AI EXTRACTION ───────────────────────────
async function extractCasesWithClaude(pdfBase64, apiKey) {
  const prompt = `Analiza este PDF de documentos médicos de Metrored / BUPA Ecuador.

El PDF contiene múltiples CASOS. Cada caso normalmente tiene 2 páginas:
  1. Estado de Cuenta de METRORED (logo Metrored arriba, datos del paciente)
  2. Autorización de Cobertura de BUPA (logo BUPA arriba, número de póliza entre paréntesis)

Para cada caso extrae:
- policy_number : El número de 6 dígitos entre paréntesis en la página BUPA.
                  Ejemplo: de "(700220)" extrae "700220". Si no hay BUPA, usa "SIN_POLIZA".
- patient_name  : Nombre completo. Campo "Asegurado:" en BUPA ó "Paciente:" en Metrored.
- pages         : Array de números de página (iniciando en 1) que pertenecen a este caso.
- document_types: Array con tipo de cada página:
                  "estado_cuenta_metrored" ó "autorizacion_cobertura_bupa" ó "otro_documento"

REGLAS:
- Si un par de páginas pertenecen al mismo paciente, agrúpalas en un solo caso.
- Ignora páginas en blanco o ilegibles.
- Devuelve ÚNICAMENTE el array JSON válido, sin texto extra, sin markdown, sin comillas de código.

Formato exacto de respuesta:
[{"policy_number":"700220","patient_name":"RAMIREZ ANCHUNDIA EDWARD FABRICIO","pages":[1,2],"document_types":["estado_cuenta_metrored","autorizacion_cobertura_bupa"]}]`;

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'pdfs-2024-09-25',
    },
    body: JSON.stringify({
      model: 'claude-opus-4-6',
      max_tokens: 4096,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'document',
            source: { type: 'base64', media_type: 'application/pdf', data: pdfBase64 }
          },
          { type: 'text', text: prompt }
        ]
      }]
    })
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Claude API error (${response.status}): ${err}`);
  }

  const data = await response.json();
  const raw  = (data.content?.[0]?.text || '').trim();

  // Strip any accidental markdown code fences
  const cleaned = raw.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();

  try {
    return JSON.parse(cleaned);
  } catch {
    throw new Error(`Respuesta de IA no es JSON válido: ${cleaned.substring(0, 200)}`);
  }
}

// ─────────────────────────── GOOGLE DRIVE ───────────────────────────
function pemToDer(pem) {
  const b64 = pem
    .replace(/-----BEGIN PRIVATE KEY-----/g, '')
    .replace(/-----END PRIVATE KEY-----/g, '')
    .replace(/\s+/g, '');
  const bin  = atob(b64);
  const buf  = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
  return buf.buffer;
}

function b64url(str) {
  return btoa(str).replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
}

function b64urlBuf(buf) {
  const bytes = new Uint8Array(buf);
  let s = '';
  bytes.forEach(b => s += String.fromCharCode(b));
  return btoa(s).replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
}

async function getGoogleDriveToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const header  = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = b64url(JSON.stringify({
    iss:   sa.client_email,
    scope: 'https://www.googleapis.com/auth/drive',
    aud:   'https://oauth2.googleapis.com/token',
    exp:   now + 3600,
    iat:   now
  }));

  const sigInput = `${header}.${payload}`;
  const keyDer   = pemToDer(sa.private_key);
  const key      = await crypto.subtle.importKey(
    'pkcs8', keyDer,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false, ['sign']
  );
  const sig = await crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5', key,
    new TextEncoder().encode(sigInput)
  );

  const jwt = `${sigInput}.${b64urlBuf(sig)}`;

  const res  = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`
  });
  const tok = await res.json();
  if (!tok.access_token) throw new Error(`Google Auth failed: ${JSON.stringify(tok)}`);
  return tok.access_token;
}

async function createDriveFolder(name, parentId, token) {
  const res = await fetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId]
    })
  });
  const d = await res.json();
  if (!d.id) throw new Error(`Error creando carpeta "${name}": ${JSON.stringify(d)}`);
  return d.id;
}

async function uploadToDrive(name, pdfBytes, folderId, token) {
  const boundary = '----MetroredBoundary314159';
  const meta     = JSON.stringify({ name, parents: [folderId] });

  const enc    = new TextEncoder();
  const part1  = enc.encode(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n`);
  const part2  = enc.encode(`--${boundary}\r\nContent-Type: application/pdf\r\n\r\n`);
  const part3  = enc.encode(`\r\n--${boundary}--`);

  const body   = new Uint8Array(part1.length + part2.length + pdfBytes.length + part3.length);
  let off = 0;
  body.set(part1, off); off += part1.length;
  body.set(part2, off); off += part2.length;
  body.set(pdfBytes, off); off += pdfBytes.length;
  body.set(part3, off);

  const res = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': `multipart/related; boundary=${boundary}`
      },
      body
    }
  );
  const d = await res.json();
  if (!d.id) throw new Error(`Error subiendo "${name}": ${JSON.stringify(d)}`);
  return d.id;
}

// ─────────────────────────── UTILITIES ───────────────────────────
function sanitize(str) {
  return str
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 80);
}

function uint8ToBase64(arr) {
  let s = '';
  const chunk = 8192;
  for (let i = 0; i < arr.length; i += chunk) {
    s += String.fromCharCode(...arr.subarray(i, i + chunk));
  }
  return btoa(s);
}
