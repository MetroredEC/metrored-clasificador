// ============================================================
//  CLASIFICADOR METRORED-BUPA  |  Cloudflare Worker v5
//  OCR desde navegador -> Worker solo crea carpetas y sube
// ============================================================
import { PDFDocument } from 'pdf-lib';

function buildHTML() {
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>Clasificador BUPA-Metrored</title>' +
'<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js"></script>' +
'<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Segoe UI,system-ui,sans-serif;background:linear-gradient(135deg,#1a56db,#7e3af2);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}.card{background:#fff;border-radius:18px;padding:40px;max-width:660px;width:100%;box-shadow:0 25px 70px rgba(0,0,0,.25)}h1{font-size:22px;color:#111827;text-align:center;margin-bottom:6px}.sub{color:#6b7280;font-size:13px;text-align:center;margin-bottom:28px}.badge{background:#ede9fe;color:#6d28d9;font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;margin-left:6px}.sec{background:#f9fafb;border-radius:10px;padding:20px;margin-bottom:16px}.sec-t{font-size:11px;font-weight:700;color:#6b7280;letter-spacing:.06em;text-transform:uppercase;margin-bottom:12px}label{display:block;font-size:13px;font-weight:600;color:#374151;margin-bottom:5px}.hint{font-weight:400;color:#9ca3af;font-size:11px;margin-left:4px}input{width:100%;padding:9px 13px;border:1.5px solid #e5e7eb;border-radius:8px;font-size:13px;outline:none}input:focus{border-color:#7c3aed}.drop{border:2px dashed #d1d5db;border-radius:10px;padding:26px;text-align:center;cursor:pointer;background:#fff;position:relative}.drop:hover{border-color:#7c3aed;background:#f5f3ff}.drop p{color:#6b7280;font-size:13px;margin-top:8px}.drop strong{color:#7c3aed}#fi{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}#fname{color:#111827;font-weight:600;font-size:13px;margin-top:6px;display:none}.btn{width:100%;padding:13px;background:linear-gradient(135deg,#1a56db,#7e3af2);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;margin-top:4px}.btn:disabled{opacity:.5;cursor:not-allowed}.prog{display:none;margin-top:20px;background:#f3f4f6;border-radius:10px;padding:18px}.bg{background:#e5e7eb;border-radius:99px;height:7px;overflow:hidden;margin:10px 0}.bar{background:linear-gradient(90deg,#1a56db,#7e3af2);height:100%;border-radius:99px;transition:width .3s;width:0%}.pmsg{font-size:13px;color:#6b7280;text-align:center}.res{display:none;margin-top:20px}.ok{background:#ecfdf5;border-left:4px solid #10b981;border-radius:6px;padding:12px 16px;margin-bottom:8px;font-size:13px;color:#065f46}.err{background:#fef2f2;border-left:4px solid #ef4444;border-radius:6px;padding:12px 16px;color:#991b1b;font-size:13px;margin-top:12px}.sum{text-align:center;background:#ede9fe;border-radius:8px;padding:12px;margin-bottom:14px;color:#5b21b6;font-weight:700;font-size:14px}</style></head><body>' +
'<div class="card"><div style="text-align:center;margin-bottom:20px"><h1>Clasificador BUPA-Metrored <span class="badge">OCR + IA</span></h1><p class="sub">Organiza escaneados masivos en SharePoint automaticamente</p></div>' +
'<div class="sec"><div class="sec-t">Configuracion</div><div style="display:grid;grid-template-columns:1fr 1fr;gap:14px"><div><label>Numero de Despacho</label><input type="text" id="despacho" placeholder="2026-04-27"></div><div><label>Carpeta SharePoint</label><input type="text" id="spFolder" placeholder="Shared Documents"></div></div></div>' +
'<div class="sec"><div class="sec-t">PDF Masivo</div><div class="drop" id="drop"><input type="file" id="fi" accept=".pdf" onchange="pickFile(this)"><p>Arrastra el PDF aqui o <strong>haz clic para seleccionar</strong></p><div id="fname"></div></div></div>' +
'<button class="btn" id="btn" onclick="run()">Clasificar y Subir a SharePoint</button>' +
'<div class="prog" id="prog"><div class="pmsg" id="pmsg">Iniciando...</div><div class="bg"><div class="bar" id="bar"></div></div></div>' +
'<div class="res" id="res"></div></div>' +
'<script>' +
'var file=null;' +
'function pickFile(inp){file=inp.files[0];var f=document.getElementById("fname");f.textContent=file.name;f.style.display="block";}' +
'var drop=document.getElementById("drop");' +
'drop.ondragover=function(e){e.preventDefault();drop.style.borderColor="#7c3aed";};' +
'drop.ondragleave=function(){drop.style.borderColor="#d1d5db";};' +
'drop.ondrop=function(e){e.preventDefault();drop.style.borderColor="#d1d5db";var f=e.dataTransfer.files[0];if(f&&f.type==="application/pdf"){file=f;var fn=document.getElementById("fname");fn.textContent=f.name;fn.style.display="block";}};' +
'function setBar(p,m){document.getElementById("bar").style.width=p+"%";document.getElementById("pmsg").textContent=m;}' +

// OCR function - called directly from browser
'async function ocrChunk(pdfBytes, ocrEndpoint, ocrKey) {' +
'var submitResp = await fetch(ocrEndpoint.replace(/\\/$/,"") + "/formrecognizer/documentModels/prebuilt-read:analyze?api-version=2023-07-31", {' +
'method:"POST", headers:{"Ocp-Apim-Subscription-Key":ocrKey,"Content-Type":"application/pdf"}, body:pdfBytes});' +
'if(!submitResp.ok) throw new Error("OCR submit: " + await submitResp.text());' +
'var opUrl = submitResp.headers.get("Operation-Location");' +
'for(var i=0;i<24;i++){' +
'await new Promise(function(r){setTimeout(r,4000);});' +
'var poll = await fetch(opUrl, {headers:{"Ocp-Apim-Subscription-Key":ocrKey}});' +
'var pd = await poll.json();' +
'if(pd.status==="succeeded"){' +
'var txt = "";' +
'var pages = (pd.analyzeResult && pd.analyzeResult.pages) ? pd.analyzeResult.pages : [];' +
'pages.forEach(function(pg){' +
'txt += "<<<PAG" + pg.pageNumber + ">>>\n";' +
'(pg.lines||[]).forEach(function(l){txt += l.content + "\n";});' +
'});' +
'return txt;' +
'}' +
'if(pd.status==="failed") throw new Error("OCR fallo");' +
'}' +
'throw new Error("OCR timeout");' +
'}' +

// Extract cases from OCR text
'function extractCases(ocrText, pageOffset){' +
'var parts = ocrText.split(/<<<PAG\\d+>>>/);' +
'var pages = parts.filter(function(p){return p.trim().length>10;});' +
'function getPatient(t){var m=t.match(/Paciente:\\s*([A-Z\xc0-\xd6\xd8-\xde][A-Z\xc0-\xd6\xd8-\xde\\s]{4,60}?)(?:\\s{2,}|\\s*Edad:|\\s*Hc:|\\n)/m);return m?m[1].trim().replace(/\\s+/g," "):null;}' +
'function getPolicy(t){var m=t.match(/\\((\\d{5,7})\\)/);return m?m[1]:"SIN_POLIZA";}' +
'function isBupa(t){return /AUTORIZACI[O\xd3]N DE COBERTURA|N[u\xfa]mero de P[o\xf3]liza/i.test(t);}' +
'function isMet(t){return /Estado de cuenta|METRORED|Paciente:/i.test(t);}' +
'var cases=[];var i=0;' +
'while(i<pages.length){' +
'var abs=pageOffset+i+1;' +
'if(isMet(pages[i])){' +
'var pat=getPatient(pages[i]);' +
'if(!pat){i++;continue;}' +
'if(i+1<pages.length&&isBupa(pages[i+1])){' +
'cases.push({policy_number:getPolicy(pages[i+1]),patient_name:pat,pages:[abs,abs+1],document_types:["estado_cuenta_metrored","autorizacion_cobertura_bupa"]});i+=2;' +
'}else{cases.push({policy_number:"SIN_POLIZA",patient_name:pat,pages:[abs],document_types:["estado_cuenta_metrored"]});i++;}' +
'}else if(isBupa(pages[i])){' +
'var am=pages[i].match(/Asegurado:\\s*([A-Z\xc0-\xd6\xd8-\xde][A-Z\xc0-\xd6\xd8-\xde\\s]{4,60}?)(?:\\s{2,}|\\s*Fecha|\\n)/m);' +
'cases.push({policy_number:getPolicy(pages[i]),patient_name:am?am[1].trim().replace(/\\s+/g," "):"PACIENTE_"+abs,pages:[abs],document_types:["autorizacion_cobertura_bupa"]});i++;' +
'}else{i++;}' +
'}return cases;' +
'}' +

'async function run(){' +
'var despacho=document.getElementById("despacho").value.trim();' +
'var spFolder=document.getElementById("spFolder").value.trim();' +
'if(!despacho||!spFolder||!file){alert("Completa todos los campos y selecciona un PDF");return;}' +
'var btn=document.getElementById("btn");btn.disabled=true;' +
'document.getElementById("prog").style.display="block";' +
'document.getElementById("res").style.display="none";' +
'try{' +
'setBar(3,"Obteniendo configuracion...");' +
'var cfgResp = await fetch("/api/config");' +
'var cfg = await cfgResp.json();' +
'if(!cfg.ocrEndpoint) throw new Error("No se pudo obtener configuracion del servidor");' +
'setBar(8,"Leyendo PDF...");' +
'var buf = await file.arrayBuffer();' +
'var fullBytes = new Uint8Array(buf);' +
'var PDFDoc = PDFLib.PDFDocument;' +
'var fullPdf = await PDFDoc.load(fullBytes);' +
'var totalPages = fullPdf.getPageCount();' +
'var chunkSize = 4;' +
'var allCases = [];' +
'var chunks = Math.ceil(totalPages/chunkSize);' +
'for(var c=0;c<chunks;c++){' +
'var startPage = c*chunkSize;' +
'var endPage = Math.min(startPage+chunkSize,totalPages);' +
'var chunkPdf = await PDFDoc.create();' +
'var indices = [];for(var p=startPage;p<endPage;p++)indices.push(p);' +
'var copied = await chunkPdf.copyPages(fullPdf,indices);' +
'copied.forEach(function(pg){chunkPdf.addPage(pg);});' +
'var chunkBytes = await chunkPdf.save();' +
'var pct = 10 + Math.round((c/chunks)*50);' +
'setBar(pct,"OCR bloque "+(c+1)+"/"+chunks+"...");' +
'var ocrText = await ocrChunk(chunkBytes, cfg.ocrEndpoint, cfg.ocrKey);' +
'var chunkCases = extractCases(ocrText, startPage);' +
'allCases = allCases.concat(chunkCases);' +
'}' +
'setBar(65,"Encontrados "+allCases.length+" casos. Subiendo a SharePoint...");' +
'var fd = new FormData();' +
'fd.append("pdf", new Blob([fullBytes],{type:"application/pdf"}), file.name);' +
'fd.append("despacho", despacho);' +
'fd.append("spFolder", spFolder);' +
'fd.append("cases", JSON.stringify(allCases));' +
'var resp = await fetch("/api/upload", {method:"POST",body:fd});' +
'var data = await resp.json();' +
'if(data.error) throw new Error(data.error);' +
'setBar(100,"Completado");' +
'var html="<div class=\\"sum\\">"+data.cases_processed+" casos procesados - Despacho: "+data.despacho+"</div>";' +
'data.cases.forEach(function(c){html+="<div class=\\"ok\\"><strong>"+c.folder+"</strong> - pags: "+c.pages.join(", ")+"</div>";});' +
'var res=document.getElementById("res");res.innerHTML=html;res.style.display="block";' +
'}catch(e){' +
'document.getElementById("res").innerHTML="<div class=\\"err\\">ERROR: "+e.message+"</div>";' +
'document.getElementById("res").style.display="block";setBar(0,"Error");' +
'}btn.disabled=false;}' +
'</script></body></html>';
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cors = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' };
    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });

    // Serve UI
    if (url.pathname === '/') return new Response(buildHTML(), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });

    // Config endpoint — returns OCR credentials to browser
    if (url.pathname === '/api/config') {
      return new Response(JSON.stringify({
        ocrEndpoint: env.DOCAI_ENDPOINT,
        ocrKey: env.DOCAI_KEY
      }), { headers: { ...cors, 'Content-Type': 'application/json' } });
    }

    // Upload endpoint — receives cases + full PDF, creates folders and uploads pages
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
  if (!allCases || allCases.length === 0) throw new Error('No se recibieron casos');

  const pdfBytes = new Uint8Array(await pdfFile.arrayBuffer());
  const pdfDoc   = await PDFDocument.load(pdfBytes);
  const total    = pdfDoc.getPageCount();

  const spToken  = await getSharePointToken(env.AZURE_TENANT_ID, env.AZURE_CLIENT_ID, env.AZURE_CLIENT_SECRET);
  const siteInfo = await getSharePointSite(spToken, env.SP_HOSTNAME, env.SP_SITE_PATH);
  const despachoPath = spFolder + '/DESPACHO-' + despacho;
  await createSharePointFolder(spToken, siteInfo.driveId, despachoPath);

  const results = [];
  for (const caso of allCases) {
    const patient  = sanitize(caso.patient_name || 'DESCONOCIDO');
    const policy   = (caso.policy_number || 'SIN_POLIZA').toString().trim();
    const folder   = policy + ' - ' + patient;
    const casePath = despachoPath + '/' + folder;
    const pages    = caso.pages || [];
    const types    = caso.document_types || [];

    await createSharePointFolder(spToken, siteInfo.driveId, casePath);

    for (let i = 0; i < pages.length; i++) {
      const idx = pages[i] - 1;
      if (idx < 0 || idx >= total) continue;
      const single = await PDFDocument.create();
      const [pg]   = await single.copyPages(pdfDoc, [idx]);
      single.addPage(pg);
      const bytes  = await single.save();
      const name   = (i + 1) + '_' + (types[i] || 'pagina_' + pages[i]) + '.pdf';
      await uploadSharePointFile(spToken, siteInfo.driveId, casePath + '/' + name, bytes);
    }

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
  if (!s.id) throw new Error('Sitio no encontrado: ' + JSON.stringify(s));
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

function sanitize(str) { return str.replace(/[<>:"/\\|?*\x00-\x1F#%{}^~\[\]`']/g,'').replace(/\s+/g,' ').trim().substring(0,60); }
