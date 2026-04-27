// ============================================================
//  CLASIFICADOR METRORED-BUPA  |  Cloudflare Worker
//  GPT-4o extrae datos de texto, Drive organiza carpetas
// ============================================================

function buildHTML() {
  return '<!DOCTYPE html>' +
'<html lang="es"><head>' +
'<meta charset="UTF-8">' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
'<title>Clasificador BUPA-Metrored</title>' +
'<style>' +
'*{box-sizing:border-box;margin:0;padding:0}' +
'body{font-family:Segoe UI,system-ui,sans-serif;background:linear-gradient(135deg,#1a56db,#7e3af2);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}' +
'.card{background:#fff;border-radius:18px;padding:40px;max-width:620px;width:100%;box-shadow:0 25px 70px rgba(0,0,0,.25)}' +
'h1{font-size:22px;color:#111827;text-align:center;margin-bottom:6px}' +
'.sub{color:#6b7280;font-size:13px;text-align:center;margin-bottom:28px}' +
'.badge{background:#ede9fe;color:#6d28d9;font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;margin-left:6px}' +
'.sec{background:#f9fafb;border-radius:10px;padding:20px;margin-bottom:16px}' +
'.sec-t{font-size:11px;font-weight:700;color:#6b7280;letter-spacing:.06em;text-transform:uppercase;margin-bottom:12px}' +
'label{display:block;font-size:13px;font-weight:600;color:#374151;margin-bottom:5px}' +
'.hint{font-weight:400;color:#9ca3af;font-size:11px;margin-left:4px}' +
'input{width:100%;padding:9px 13px;border:1.5px solid #e5e7eb;border-radius:8px;font-size:13px;outline:none}' +
'input:focus{border-color:#7c3aed}' +
'.row{display:grid;grid-template-columns:1fr 1fr;gap:14px}' +
'.drop{border:2px dashed #d1d5db;border-radius:10px;padding:26px;text-align:center;cursor:pointer;background:#fff;position:relative}' +
'.drop:hover{border-color:#7c3aed;background:#f5f3ff}' +
'.drop p{color:#6b7280;font-size:13px;margin-top:8px}' +
'.drop strong{color:#7c3aed}' +
'#fi{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}' +
'#fname{color:#111827;font-weight:600;font-size:13px;margin-top:6px;display:none}' +
'.btn{width:100%;padding:13px;background:linear-gradient(135deg,#1a56db,#7e3af2);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;margin-top:4px}' +
'.btn:disabled{opacity:.5;cursor:not-allowed}' +
'.prog{display:none;margin-top:20px;background:#f3f4f6;border-radius:10px;padding:18px}' +
'.bg{background:#e5e7eb;border-radius:99px;height:7px;overflow:hidden;margin:10px 0}' +
'.bar{background:linear-gradient(90deg,#1a56db,#7e3af2);height:100%;border-radius:99px;transition:width .4s;width:0%}' +
'.pmsg{font-size:13px;color:#6b7280;text-align:center}' +
'.res{display:none;margin-top:20px}' +
'.ok{background:#ecfdf5;border-left:4px solid #10b981;border-radius:6px;padding:12px 16px;margin-bottom:8px;font-size:13px;color:#065f46}' +
'.err{background:#fef2f2;border-left:4px solid #ef4444;border-radius:6px;padding:12px 16px;color:#991b1b;font-size:13px;margin-top:12px}' +
'.sum{text-align:center;background:#ede9fe;border-radius:8px;padding:12px;margin-bottom:14px;color:#5b21b6;font-weight:700;font-size:14px}' +
'</style></head><body>' +
'<div class="card">' +
'<div style="text-align:center;margin-bottom:20px">' +
'<h1>Clasificador BUPA-Metrored <span class="badge">GPT-4o</span></h1>' +
'<p class="sub">Organiza escaneados masivos en Google Drive automaticamente</p>' +
'</div>' +
'<div class="sec">' +
'<div class="sec-t">Informacion del Lote</div>' +
'<div class="row">' +
'<div><label>Numero de Despacho <span class="hint">ej: 2026-04-08</span></label>' +
'<input type="text" id="despacho" placeholder="2026-04-08"></div>' +
'<div><label>Carpeta Drive Destino <span class="hint">URL o ID</span></label>' +
'<input type="text" id="driveUrl" placeholder="https://drive.google.com/drive/folders/..."></div>' +
'</div></div>' +
'<div class="sec">' +
'<div class="sec-t">PDF Masivo</div>' +
'<div class="drop" id="drop">' +
'<input type="file" id="fi" accept=".pdf" onchange="pickFile(this)">' +
'<p>Arrastra el PDF aqui o <strong>haz clic para seleccionar</strong></p>' +
'<div id="fname"></div>' +
'</div></div>' +
'<button class="btn" id="btn" onclick="run()">Clasificar y Subir a Drive</button>' +
'<div class="prog" id="prog"><div class="pmsg" id="pmsg">Iniciando...</div><div class="bg"><div class="bar" id="bar"></div></div></div>' +
'<div class="res" id="res"></div>' +
'</div>' +
'<script>' +
'var file=null;' +
'function pickFile(inp){file=inp.files[0];var f=document.getElementById("fname");f.textContent=file.name;f.style.display="block";}' +
'var drop=document.getElementById("drop");' +
'drop.ondragover=function(e){e.preventDefault();drop.style.borderColor="#7c3aed";};' +
'drop.ondragleave=function(){drop.style.borderColor="#d1d5db";};' +
'drop.ondrop=function(e){e.preventDefault();drop.style.borderColor="#d1d5db";var f=e.dataTransfer.files[0];if(f&&f.type==="application/pdf"){file=f;var fn=document.getElementById("fname");fn.textContent=f.name;fn.style.display="block";}};' +
'function setBar(p,m){document.getElementById("bar").style.width=p+"%";document.getElementById("pmsg").textContent=m;}' +
'function run(){' +
'var despacho=document.getElementById("despacho").value.trim();' +
'var driveRaw=document.getElementById("driveUrl").value.trim();' +
'if(!despacho){alert("Ingresa el numero de despacho");return;}' +
'if(!driveRaw){alert("Ingresa la URL de la carpeta de Drive");return;}' +
'if(!file){alert("Selecciona un archivo PDF");return;}' +
'var m=driveRaw.match(/\\/folders\\/([a-zA-Z0-9_\\-]+)/);' +
'var folderId=m?m[1]:driveRaw;' +
'var btn=document.getElementById("btn");' +
'btn.disabled=true;' +
'document.getElementById("prog").style.display="block";' +
'document.getElementById("res").style.display="none";' +
'setBar(8,"Leyendo PDF...");' +
'var fd=new FormData();' +
'fd.append("pdf",file);' +
'fd.append("despacho",despacho);' +
'fd.append("driveFolderId",folderId);' +
'setBar(20,"Extrayendo datos con IA (30-60 seg)...");' +
'fetch("/api/process",{method:"POST",body:fd}).then(function(resp){' +
'setBar(75,"Creando carpetas y subiendo a Drive...");' +
'return resp.json().then(function(data){' +
'if(!resp.ok||data.error)throw new Error(data.error||"Error desconocido");' +
'setBar(100,"Proceso completado");' +
'var html="<div class=\\"sum\\">"+data.cases_processed+" casos procesados - Despacho: "+data.despacho+"</div>";' +
'data.cases.forEach(function(c){html+="<div class=\\"ok\\"><strong>"+c.folder+"</strong></div>";});' +
'var res=document.getElementById("res");res.innerHTML=html;res.style.display="block";' +
'});' +
'}).catch(function(e){' +
'document.getElementById("res").innerHTML="<div class=\\"err\\">ERROR: "+e.message+"</div>";' +
'document.getElementById("res").style.display="block";' +
'setBar(0,"Error");' +
'}).finally(function(){btn.disabled=false;});' +
'}' +
'</script></body></html>';
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };
    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });
    if (url.pathname === '/') return new Response(buildHTML(), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    if (url.pathname === '/api/process' && request.method === 'POST') {
      try {
        const result = await handleProcess(request, env);
        return new Response(JSON.stringify(result), { headers: { ...cors, 'Content-Type': 'application/json' } });
      } catch (err) {
        return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { ...cors, 'Content-Type': 'application/json' } });
      }
    }
    return new Response('Not found', { status: 404 });
  }
};

async function handleProcess(request, env) {
  const form = await request.formData();
  const pdfFile = form.get('pdf');
  const despachoNum = form.get('despacho');
  const rootFolderId = form.get('driveFolderId');
  if (!pdfFile || !despachoNum || !rootFolderId) throw new Error('Faltan campos requeridos');

  const pdfBytes = new Uint8Array(await pdfFile.arrayBuffer());
  const pdfBase64 = uint8ToBase64(pdfBytes);

  const cases = await extractCasesWithGPT(pdfBase64, env.OPENAI_API_KEY);
  if (!Array.isArray(cases) || cases.length === 0) throw new Error('No se pudieron extraer casos del PDF');

  const serviceAccount = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT);
  const driveToken = await getGoogleDriveToken(serviceAccount);

  const despachoFolderId = await createDriveFolder('DESPACHO-' + despachoNum, rootFolderId, driveToken);

  const results = [];
  for (const caso of cases) {
    const patientClean = sanitize(caso.patient_name || 'PACIENTE_DESCONOCIDO');
    const policyNum = (caso.policy_number || 'SIN_POLIZA').toString().trim();
    const folderName = policyNum + ' - ' + patientClean;
    await createDriveFolder(folderName, despachoFolderId, driveToken);
    results.push({ folder: folderName, pages: caso.pages || [], status: 'OK' });
  }

  return { success: true, despacho: despachoNum, cases_processed: results.length, cases: results };
}

async function extractCasesWithGPT(pdfBase64, apiKey) {
  const bin = atob(pdfBase64);
  const readable = bin.replace(/[^\x20-\x7E\n\r\t]/g, ' ').replace(/ {4,}/g, ' ').replace(/\n{3,}/g, '\n\n');
  const pdfText = readable.substring(0, 14000);

  const prompt = 'Analiza el siguiente texto extraido de un PDF de documentos medicos Metrored/BUPA Ecuador.\n\n' +
    'Cada caso tiene 2 documentos: 1) Estado de Cuenta METRORED y 2) Autorizacion BUPA.\n' +
    'La pagina BUPA contiene el numero de poliza entre parentesis como "(700220)".\n\n' +
    'Extrae todos los casos. Devuelve UNICAMENTE un array JSON valido sin texto extra ni markdown.\n\n' +
    'Campos por caso:\n' +
    '- policy_number: numero 6 digitos entre parentesis de BUPA. Si no hay, "SIN_POLIZA"\n' +
    '- patient_name: nombre completo MAYUSCULAS del campo Asegurado o Paciente\n' +
    '- pages: array de numeros de pagina del caso (empezando en 1)\n' +
    '- document_types: array con "estado_cuenta_metrored" o "autorizacion_cobertura_bupa"\n\n' +
    'Ejemplo: [{"policy_number":"700220","patient_name":"RAMIREZ ANCHUNDIA EDWARD","pages":[1,2],"document_types":["estado_cuenta_metrored","autorizacion_cobertura_bupa"]}]\n\n' +
    'TEXTO DEL PDF:\n' + pdfText;

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
    body: JSON.stringify({ model: 'gpt-4o', max_tokens: 2048, messages: [{ role: 'user', content: prompt }] })
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error('OpenAI API error (' + response.status + '): ' + err);
  }

  const data = await response.json();
  const raw = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) ? data.choices[0].message.content.trim() : '';
  const cleaned = raw.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();

  try { return JSON.parse(cleaned); }
  catch (e) { throw new Error('Respuesta de IA no es JSON valido: ' + cleaned.substring(0, 300)); }
}

function pemToDer(pem) {
  const b64 = pem.replace(/-----BEGIN PRIVATE KEY-----/g, '').replace(/-----END PRIVATE KEY-----/g, '').replace(/\s+/g, '');
  const bin = atob(b64);
  const buf = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
  return buf.buffer;
}

function b64url(str) { return btoa(str).replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_'); }
function b64urlBuf(buf) { const b = new Uint8Array(buf); let s = ''; b.forEach(function(x) { s += String.fromCharCode(x); }); return btoa(s).replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_'); }

async function getGoogleDriveToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const header = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = b64url(JSON.stringify({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/drive', aud: 'https://oauth2.googleapis.com/token', exp: now + 3600, iat: now }));
  const sigInput = header + '.' + payload;
  const keyDer = pemToDer(sa.private_key);
  const key = await crypto.subtle.importKey('pkcs8', keyDer, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['sign']);
  const sig = await crypto.subtle.sign('RSASSA-PKCS1-v1_5', key, new TextEncoder().encode(sigInput));
  const jwt = sigInput + '.' + b64urlBuf(sig);
  const res = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + jwt });
  const tok = await res.json();
  if (!tok.access_token) throw new Error('Google Auth failed: ' + JSON.stringify(tok));
  return tok.access_token;
}

async function createDriveFolder(name, parentId, token) {
  const res = await fetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: name, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] })
  });
  const d = await res.json();
  if (!d.id) throw new Error('Error creando carpeta "' + name + '": ' + JSON.stringify(d));
  return d.id;
}

async function uploadToDrive(name, fileBytes, folderId, token, mimeType) {
  mimeType = mimeType || 'text/plain';
  const boundary = 'MetroredBoundary314159';
  const meta = JSON.stringify({ name: name, parents: [folderId] });
  const enc = new TextEncoder();
  const p1 = enc.encode('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n');
  const p2 = enc.encode('--' + boundary + '\r\nContent-Type: ' + mimeType + '\r\n\r\n');
  const p3 = enc.encode('\r\n--' + boundary + '--');
  const body = new Uint8Array(p1.length + p2.length + fileBytes.length + p3.length);
  let off = 0; body.set(p1, off); off += p1.length; body.set(p2, off); off += p2.length; body.set(fileBytes, off); off += fileBytes.length; body.set(p3, off);
  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'multipart/related; boundary=' + boundary },
    body: body
  });
  const d = await res.json();
  if (!d.id) throw new Error('Error subiendo "' + name + '": ' + JSON.stringify(d));
  return d.id;
}

function sanitize(str) { return str.replace(/[<>:"/\\|?*\x00-\x1F]/g, '').replace(/\s+/g, ' ').trim().substring(0, 80); }
function uint8ToBase64(arr) { let s = ''; const c = 8192; for (let i = 0; i < arr.length; i += c) s += String.fromCharCode.apply(null, arr.subarray(i, i + c)); return btoa(s); }
