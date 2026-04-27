// ============================================================
//  CLASIFICADOR METRORED-BUPA  |  Cloudflare Worker
//  Procesa chunks de 4 paginas desde el navegador
// ============================================================
import { PDFDocument } from 'pdf-lib';

function buildHTML() {
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>Clasificador BUPA-Metrored</title>' +
'<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js"></script>' +
'<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Segoe UI,system-ui,sans-serif;background:linear-gradient(135deg,#1a56db,#7e3af2);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}.card{background:#fff;border-radius:18px;padding:40px;max-width:640px;width:100%;box-shadow:0 25px 70px rgba(0,0,0,.25)}h1{font-size:22px;color:#111827;text-align:center;margin-bottom:6px}.sub{color:#6b7280;font-size:13px;text-align:center;margin-bottom:28px}.badge{background:#ede9fe;color:#6d28d9;font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;margin-left:6px}.sec{background:#f9fafb;border-radius:10px;padding:20px;margin-bottom:16px}.sec-t{font-size:11px;font-weight:700;color:#6b7280;letter-spacing:.06em;text-transform:uppercase;margin-bottom:12px}label{display:block;font-size:13px;font-weight:600;color:#374151;margin-bottom:5px}.hint{font-weight:400;color:#9ca3af;font-size:11px;margin-left:4px}input{width:100%;padding:9px 13px;border:1.5px solid #e5e7eb;border-radius:8px;font-size:13px;outline:none}input:focus{border-color:#7c3aed}.drop{border:2px dashed #d1d5db;border-radius:10px;padding:26px;text-align:center;cursor:pointer;background:#fff;position:relative}.drop:hover{border-color:#7c3aed;background:#f5f3ff}.drop p{color:#6b7280;font-size:13px;margin-top:8px}.drop strong{color:#7c3aed}#fi{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}#fname{color:#111827;font-weight:600;font-size:13px;margin-top:6px;display:none}.btn{width:100%;padding:13px;background:linear-gradient(135deg,#1a56db,#7e3af2);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;margin-top:4px}.btn:disabled{opacity:.5;cursor:not-allowed}.prog{display:none;margin-top:20px;background:#f3f4f6;border-radius:10px;padding:18px}.bg{background:#e5e7eb;border-radius:99px;height:7px;overflow:hidden;margin:10px 0}.bar{background:linear-gradient(90deg,#1a56db,#7e3af2);height:100%;border-radius:99px;transition:width .4s;width:0%}.pmsg{font-size:13px;color:#6b7280;text-align:center}.res{display:none;margin-top:20px}.ok{background:#ecfdf5;border-left:4px solid #10b981;border-radius:6px;padding:12px 16px;margin-bottom:8px;font-size:13px;color:#065f46}.err{background:#fef2f2;border-left:4px solid #ef4444;border-radius:6px;padding:12px 16px;color:#991b1b;font-size:13px;margin-top:12px}.sum{text-align:center;background:#ede9fe;border-radius:8px;padding:12px;margin-bottom:14px;color:#5b21b6;font-weight:700;font-size:14px}</style></head><body>' +
'<div class="card"><div style="text-align:center;margin-bottom:20px"><h1>Clasificador BUPA-Metrored <span class="badge">OCR + GPT-4o</span></h1><p class="sub">Organiza escaneados masivos en SharePoint automaticamente</p></div>' +
'<div class="sec"><div class="sec-t">Informacion del Lote</div><div style="display:grid;grid-template-columns:1fr 1fr;gap:14px"><div><label>Numero de Despacho <span class="hint">ej: 2026-04-27</span></label><input type="text" id="despacho" placeholder="2026-04-27"></div><div><label>Carpeta SharePoint <span class="hint">nombre exacto</span></label><input type="text" id="spFolder" placeholder="Shared Documents"></div></div></div>' +
'<div class="sec"><div class="sec-t">PDF Masivo Escaneado</div><div class="drop" id="drop"><input type="file" id="fi" accept=".pdf" onchange="pickFile(this)"><p>Arrastra el PDF aqui o <strong>haz clic para seleccionar</strong></p><div id="fname"></div></div></div>' +
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
'function toBase64(buf){var b=new Uint8Array(buf),s="",c=8192;for(var i=0;i<b.length;i+=c)s+=String.fromCharCode.apply(null,b.subarray(i,i+c));return btoa(s);}' +
'async function run(){' +
'var despacho=document.getElementById("despacho").value.trim();' +
'var spFolder=document.getElementById("spFolder").value.trim();' +
'if(!despacho){alert("Ingresa el numero de despacho");return;}' +
'if(!spFolder){alert("Ingresa el nombre de la carpeta en SharePoint");return;}' +
'if(!file){alert("Selecciona un archivo PDF");return;}' +
'var btn=document.getElementById("btn");btn.disabled=true;' +
'document.getElementById("prog").style.display="block";' +
'document.getElementById("res").style.display="none";' +
'try{' +
'setBar(5,"Leyendo PDF...");' +
'var arrayBuf=await file.arrayBuffer();' +
'var pdfBytes=new Uint8Array(arrayBuf);' +
'var PDFDoc=PDFLib.PDFDocument;' +
'var fullPdf=await PDFDoc.load(pdfBytes);' +
'var totalPages=fullPdf.getPageCount();' +
'setBar(10,"PDF cargado: "+totalPages+" paginas. Procesando en bloques...");' +
'var chunkSize=4;' +
'var allCases=[];' +
'var chunks=Math.ceil(totalPages/chunkSize);' +
'for(var c=0;c<chunks;c++){' +
'var startPage=c*chunkSize;' +
'var endPage=Math.min(startPage+chunkSize,totalPages);' +
'var chunkPdf=await PDFDoc.create();' +
'var pageIndices=[];' +
'for(var p=startPage;p<endPage;p++)pageIndices.push(p);' +
'var copiedPages=await chunkPdf.copyPages(fullPdf,pageIndices);' +
'copiedPages.forEach(function(pg){chunkPdf.addPage(pg);});' +
'var chunkBytes=await chunkPdf.save();' +
'var pct=10+Math.round((c/chunks)*60);' +
'setBar(pct,"Bloque "+(c+1)+"/"+chunks+": OCR e identificacion...");' +
'var fd=new FormData();' +
'fd.append("pdf",new Blob([chunkBytes],{type:"application/pdf"}),file.name);' +
'fd.append("despacho",despacho);' +
'fd.append("spFolder",spFolder);' +
'fd.append("pageOffset",String(startPage));' +
'fd.append("mode","extract_only");' +
'var resp=await fetch("/api/process",{method:"POST",body:fd});' +
'var data=await resp.json();' +
'if(data.error)throw new Error("Bloque "+(c+1)+": "+data.error);' +
'if(data.cases)allCases=allCases.concat(data.cases);' +
'}' +
'setBar(75,"Creando carpetas y subiendo PDFs a SharePoint...");' +
'var fd2=new FormData();' +
'fd2.append("pdf",new Blob([pdfBytes],{type:"application/pdf"}),file.name);' +
'fd2.append("despacho",despacho);' +
'fd2.append("spFolder",spFolder);' +
'fd2.append("cases",JSON.stringify(allCases));' +
'fd2.append("mode","upload_only");' +
'var resp2=await fetch("/api/process",{method:"POST",body:fd2});' +
'var data2=await resp2.json();' +
'if(data2.error)throw new Error(data2.error);' +
'setBar(100,"Proceso completado");' +
'var html="<div class=\\"sum\\">"+data2.cases_processed+" casos procesados - Despacho: "+data2.despacho+"</div>";' +
'data2.cases.forEach(function(c){html+="<div class=\\"ok\\"><strong>"+c.folder+"</strong> - paginas: "+c.pages.join(", ")+"</div>";});' +
'var res=document.getElementById("res");res.innerHTML=html;res.style.display="block";' +
'}catch(e){' +
'document.getElementById("res").innerHTML="<div class=\\"err\\">ERROR: "+e.message+"</div>";' +
'document.getElementById("res").style.display="block";' +
'setBar(0,"Error");' +
'}' +
'btn.disabled=false;' +
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
  const form     = await request.formData();
  const mode     = form.get('mode') || 'full';
  const pdfFile  = form.get('pdf');
  const despacho = form.get('despacho');
  const spFolder = form.get('spFolder');

  const pdfBytes = new Uint8Array(await pdfFile.arrayBuffer());

  // MODE 1: extract_only — OCR a chunk and return cases (no SharePoint)
  if (mode === 'extract_only') {
    const pageOffset = parseInt(form.get('pageOffset') || '0');
    const ocrText = await extractTextWithOCR(pdfBytes, env.DOCAI_ENDPOINT, env.DOCAI_KEY);
    const cases   = extractCasesFromOCR(ocrText, pageOffset);
    return { success: true, cases };
  }

  // MODE 2: upload_only — receive all cases, create folders, split and upload PDF pages
  if (mode === 'upload_only') {
    const casesJson = form.get('cases');
    const allCases  = JSON.parse(casesJson);
    if (!allCases || allCases.length === 0) throw new Error('No se recibieron casos para subir');

    const spToken  = await getSharePointToken(env.AZURE_TENANT_ID, env.AZURE_CLIENT_ID, env.AZURE_CLIENT_SECRET);
    const siteInfo = await getSharePointSite(spToken, env.SP_HOSTNAME, env.SP_SITE_PATH);
    const despachoPath = spFolder + '/DESPACHO-' + despacho;
    await createSharePointFolder(spToken, siteInfo.driveId, despachoPath);

    const pdfDoc = await PDFDocument.load(pdfBytes);
    const totalPages = pdfDoc.getPageCount();
    const results = [];

    for (const caso of allCases) {
      const patientClean = sanitize(caso.patient_name || 'PACIENTE_DESCONOCIDO');
      const policyNum    = (caso.policy_number || 'SIN_POLIZA').toString().trim();
      const folderName   = policyNum + ' - ' + patientClean;
      const casePath     = despachoPath + '/' + folderName;
      const pages        = caso.pages || [];
      const docTypes     = caso.document_types || [];

      await createSharePointFolder(spToken, siteInfo.driveId, casePath);

      for (let i = 0; i < pages.length; i++) {
        const pageIndex = pages[i] - 1;
        if (pageIndex < 0 || pageIndex >= totalPages) continue;
        const singleDoc = await PDFDocument.create();
        const [copied]  = await singleDoc.copyPages(pdfDoc, [pageIndex]);
        singleDoc.addPage(copied);
        const singleBytes = await singleDoc.save();
        const docType  = docTypes[i] || ('pagina_' + pages[i]);
        const fileName = (i + 1) + '_' + docType + '.pdf';
        await uploadSharePointFile(spToken, siteInfo.driveId, casePath + '/' + fileName, singleBytes);
      }

      results.push({ folder: folderName, pages, status: 'OK' });
    }

    return { success: true, despacho, cases_processed: results.length, cases: results };
  }

  throw new Error('Modo no reconocido: ' + mode);
}

// ─── OCR ─────────────────────────────────────────────────────
async function extractTextWithOCR(pdfBytes, endpoint, key) {
  const submitUrl = endpoint.replace(/\/$/, '') + '/formrecognizer/documentModels/prebuilt-read:analyze?api-version=2023-07-31';
  const submitRes = await fetch(submitUrl, {
    method: 'POST',
    headers: { 'Ocp-Apim-Subscription-Key': key, 'Content-Type': 'application/pdf' },
    body: pdfBytes
  });
  if (!submitRes.ok) throw new Error('OCR submit error (' + submitRes.status + '): ' + await submitRes.text());

  const operationUrl = submitRes.headers.get('Operation-Location');
  if (!operationUrl) throw new Error('OCR no devolvio Operation-Location');

  for (let i = 0; i < 18; i++) {
    await new Promise(function(r){ setTimeout(r, 5000); });
    const pollRes  = await fetch(operationUrl, { headers: { 'Ocp-Apim-Subscription-Key': key } });
    const pollData = await pollRes.json();
    if (pollData.status === 'succeeded') {
      const pages = (pollData.analyzeResult && pollData.analyzeResult.pages) ? pollData.analyzeResult.pages : [];
      let fullText = '';
      for (const page of pages) {
        fullText += '\n<<<PAG' + page.pageNumber + '>>>\n';
        for (const line of (page.lines || [])) fullText += line.content + '\n';
      }
      return fullText.trim();
    }
    if (pollData.status === 'failed') throw new Error('OCR fallo: ' + JSON.stringify(pollData.error));
  }
  throw new Error('OCR timeout');
}

// ─── EXTRACT CASES FROM OCR TEXT ─────────────────────────────
function extractCasesFromOCR(ocrText, pageOffset) {
  const pageMarkerRegex = /<<<PAG\d+>>>/g;
  const parts = ocrText.split(pageMarkerRegex);
  const pages = parts.filter(function(p){ return p.trim().length > 10; });

  function extractPatient(text) {
    var m = text.match(/Paciente:\s*([A-ZÑÁÉÍÓÚÜ][A-ZÑÁÉÍÓÚÜ\s]{4,60}?)(?:\s{2,}|\s*Edad:|\s*Hc:|\n)/m);
    return m ? m[1].trim().replace(/\s+/g, ' ') : null;
  }

  function extractPolicy(text) {
    var m = text.match(/\((\d{5,7})\)/);
    return m ? m[1] : null;
  }

  function isBupa(text) {
    return /AUTORIZACI[OÓ]N DE COBERTURA|Numero de P[oó]liza|N[uú]mero de P[oó]liza/i.test(text);
  }

  function isMetrored(text) {
    return /Estado de cuenta|METRORED|Paciente:/i.test(text);
  }

  var cases = [];
  var i = 0;

  while (i < pages.length) {
    var absPage = pageOffset + i + 1;
    if (isMetrored(pages[i])) {
      var patient = extractPatient(pages[i]);
      if (!patient) { i++; continue; }
      if (i + 1 < pages.length && isBupa(pages[i+1])) {
        var policy = extractPolicy(pages[i+1]) || 'SIN_POLIZA';
        cases.push({
          policy_number: policy,
          patient_name: patient,
          pages: [absPage, absPage + 1],
          document_types: ['estado_cuenta_metrored', 'autorizacion_cobertura_bupa']
        });
        i += 2;
      } else {
        cases.push({
          policy_number: 'SIN_POLIZA',
          patient_name: patient,
          pages: [absPage],
          document_types: ['estado_cuenta_metrored']
        });
        i++;
      }
    } else if (isBupa(pages[i])) {
      var policy2 = extractPolicy(pages[i]) || 'SIN_POLIZA';
      var aseg = pages[i].match(/Asegurado:\s*([A-ZÑÁÉÍÓÚÜ][A-ZÑÁÉÍÓÚÜ\s]{4,60}?)(?:\s{2,}|\s*Fecha|\n)/m);
      var patient2 = aseg ? aseg[1].trim().replace(/\s+/g, ' ') : 'PACIENTE_PAG_' + absPage;
      cases.push({
        policy_number: policy2,
        patient_name: patient2,
        pages: [absPage],
        document_types: ['autorizacion_cobertura_bupa']
      });
      i++;
    } else {
      i++;
    }
  }

  return cases;
}

// ─── SHAREPOINT ───────────────────────────────────────────────
async function getSharePointToken(tenantId, clientId, clientSecret) {
  const body = 'grant_type=client_credentials&client_id=' + encodeURIComponent(clientId) + '&client_secret=' + encodeURIComponent(clientSecret) + '&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default';
  const res = await fetch('https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  const tok = await res.json();
  if (!tok.access_token) throw new Error('Azure auth failed: ' + JSON.stringify(tok));
  return tok.access_token;
}

async function getSharePointSite(token, hostname, sitePath) {
  const siteRes = await fetch('https://graph.microsoft.com/v1.0/sites/' + hostname + ':/' + sitePath, { headers: { 'Authorization': 'Bearer ' + token } });
  const site = await siteRes.json();
  if (!site.id) throw new Error('Sitio SharePoint no encontrado: ' + JSON.stringify(site));
  const driveRes = await fetch('https://graph.microsoft.com/v1.0/sites/' + site.id + '/drive', { headers: { 'Authorization': 'Bearer ' + token } });
  const drive = await driveRes.json();
  if (!drive.id) throw new Error('Drive no encontrado: ' + JSON.stringify(drive));
  return { siteId: site.id, driveId: drive.id };
}

async function createSharePointFolder(token, driveId, folderPath) {
  const parts = folderPath.split('/').filter(Boolean);
  let parentId = 'root';
  for (const part of parts) {
    const res = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + parentId + '/children', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: part, folder: {} })
    });
    const d = await res.json();
    if (d.error && d.error.code === 'nameAlreadyExists') {
      const ex = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + parentId + ':/' + encodeURIComponent(part), { headers: { 'Authorization': 'Bearer ' + token } });
      const exd = await ex.json();
      parentId = exd.id;
    } else if (!d.id) {
      throw new Error('Error creando carpeta "' + part + '": ' + JSON.stringify(d));
    } else {
      parentId = d.id;
    }
  }
  return parentId;
}

async function uploadSharePointFile(token, driveId, filePath, fileBytes) {
  const parts    = filePath.split('/').filter(Boolean);
  const fileName = parts.pop();
  const folderPath = parts.join('/');
  const folderRes = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/root:/' + folderPath, { headers: { 'Authorization': 'Bearer ' + token } });
  const folder = await folderRes.json();
  if (!folder.id) throw new Error('Carpeta no encontrada: ' + folderPath);
  const uploadUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + folder.id + ':/' + encodeURIComponent(fileName) + ':/content';
  const res = await fetch(uploadUrl, { method: 'PUT', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/pdf' }, body: fileBytes });
  const d = await res.json();
  if (!d.id) throw new Error('Error subiendo "' + fileName + '": ' + JSON.stringify(d));
  return d.id;
}

function sanitize(str) { return str.replace(/[<>:"/\\|?*\x00-\x1F#%{}^~\[\]`']/g, '').replace(/\s+/g, ' ').trim().substring(0, 60); }
function uint8ToBase64(arr) { let s=''; const c=8192; for(let i=0;i<arr.length;i+=c)s+=String.fromCharCode.apply(null,arr.subarray(i,i+c)); return btoa(s); }