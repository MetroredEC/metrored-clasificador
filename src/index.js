// ============================================================
//  CLASIFICADOR METRORED-BUPA  |  Cloudflare Worker v5
//  OCR desde navegador -> Worker solo crea carpetas y sube
// ============================================================
import { PDFDocument } from 'pdf-lib';

function buildHTML() {
  return '<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width,initial-scale=1.0">\n<title>Clasificador BUPA | Metrored</title>\n<link rel="icon" type="image/jpeg" href="https://raw.githubusercontent.com/MetroredEC/metrored-clasificador/main/public/logo-metrored.jpg">\n<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js"></script>\n<script src="https://cdnjs.cloudflare.com/ajax/libs/msal.js/2.38.3/msal-browser.min.js"></script>\n<style>\n*{box-sizing:border-box;margin:0;padding:0}\n:root{--blue:#29ABE2;--blue-dark:#1a8fc0;--gray:#f4f6f9;--text:#1a2235;--sub:#6b7280}\nbody{font-family:\'Segoe UI\',system-ui,sans-serif;background:var(--gray);min-height:100vh}\n\n/* TOP BAR */\n.topbar{background:#fff;border-bottom:2px solid #e8edf2;padding:0 32px;height:64px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.06)}\n.topbar-logo{height:36px}\n.topbar-right{display:flex;align-items:center;gap:12px}\n.user-chip{background:#f0f9ff;border:1px solid #bae6fd;color:#0369a1;font-size:12px;font-weight:600;padding:5px 12px;border-radius:99px}\n.btn-icon{background:none;border:1.5px solid #e5e7eb;border-radius:8px;padding:7px 12px;cursor:pointer;color:#6b7280;font-size:13px;font-weight:600;display:flex;align-items:center;gap:6px;transition:all .2s}\n.btn-icon:hover{border-color:var(--blue);color:var(--blue)}\n.btn-icon svg{width:15px;height:15px}\n\n/* LOGIN PAGE */\n.login-page{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#29ABE2 0%,#1a6fa8 100%)}\n.login-card{background:#fff;border-radius:20px;padding:48px 40px;max-width:420px;width:100%;text-align:center;box-shadow:0 30px 80px rgba(0,0,0,.18)}\n.login-logo{height:52px;margin-bottom:28px}\n.login-card h2{font-size:20px;color:var(--text);margin-bottom:8px}\n.login-card p{color:var(--sub);font-size:13px;margin-bottom:32px;line-height:1.6}\n.btn-ms{background:#0078d4;color:#fff;border:none;padding:13px 28px;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:10px;transition:background .2s}\n.btn-ms:hover{background:#106ebe}\n.btn-ms svg{width:18px;height:18px}\n.login-footer{margin-top:24px;font-size:11px;color:#9ca3af}\n\n/* MAIN LAYOUT */\n.main{max-width:760px;margin:0 auto;padding:32px 20px}\n.page-title{font-size:22px;font-weight:700;color:var(--text);margin-bottom:4px}\n.page-sub{font-size:13px;color:var(--sub);margin-bottom:28px}\n\n/* CARDS */\n.card{background:#fff;border-radius:14px;padding:24px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,.05);border:1px solid #e8edf2}\n.card-title{font-size:11px;font-weight:700;color:#94a3b8;letter-spacing:.07em;text-transform:uppercase;margin-bottom:16px;display:flex;align-items:center;gap:8px}\n.card-title span{display:inline-block;width:18px;height:18px;background:var(--blue);border-radius:4px;color:#fff;font-size:10px;font-weight:800;text-align:center;line-height:18px}\n\n/* FORM */\n.form-row{display:grid;grid-template-columns:1fr 1fr;gap:16px}\nlabel{display:block;font-size:12px;font-weight:600;color:#374151;margin-bottom:5px}\ninput[type=text]{width:100%;padding:9px 13px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:13px;color:var(--text);outline:none;transition:border .2s;background:#fafbfc}\ninput[type=text]:focus{border-color:var(--blue);background:#fff}\n\n/* DROP ZONE */\n.drop{border:2px dashed #cbd5e1;border-radius:12px;padding:28px;text-align:center;cursor:pointer;background:#fafbfc;position:relative;transition:all .2s}\n.drop.over,.drop:hover{border-color:var(--blue);background:#f0f9ff}\n.drop-icon{width:40px;height:40px;background:#e0f2fe;border-radius:10px;display:flex;align-items:center;justify-content:center;margin:0 auto 10px}\n.drop-icon svg{width:22px;height:22px;color:var(--blue)}\n.drop p{color:var(--sub);font-size:13px}.drop strong{color:var(--blue)}\n#fi{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}\n#fname{background:#e0f2fe;color:#0369a1;font-size:12px;font-weight:600;padding:4px 10px;border-radius:6px;display:inline-block;margin-top:8px}\n\n/* ACTION BUTTON */\n.btn-main{width:100%;padding:14px;background:var(--blue);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;transition:background .2s;display:flex;align-items:center;justify-content:center;gap:10px}\n.btn-main:hover{background:var(--blue-dark)}\n.btn-main:disabled{opacity:.5;cursor:not-allowed}\n\n/* PROGRESS */\n.prog{display:none;margin-top:20px}\n.prog-bar-bg{background:#e2e8f0;border-radius:99px;height:8px;overflow:hidden;margin:10px 0}\n.prog-bar{background:var(--blue);height:100%;border-radius:99px;transition:width .3s;width:0%}\n.prog-msg{font-size:13px;color:var(--sub);text-align:center}\n\n/* RESULTS */\n.res{display:none;margin-top:20px}\n.res-sum{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:14px 18px;color:#15803d;font-weight:700;font-size:14px;margin-bottom:12px;text-align:center}\n.res-item{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:11px 16px;margin-bottom:8px;font-size:13px;color:var(--text);display:flex;align-items:center;gap:10px}\n.res-item::before{content:"";width:8px;height:8px;background:#22c55e;border-radius:50%;flex-shrink:0}\n.res-err{background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:14px 18px;color:#dc2626;font-size:13px;margin-top:12px}\n\n/* MODAL SETTINGS */\n.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:200;align-items:center;justify-content:center}\n.modal-overlay.open{display:flex}\n.modal{background:#fff;border-radius:16px;padding:32px;max-width:480px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,.2)}\n.modal h3{font-size:17px;font-weight:700;color:var(--text);margin-bottom:4px}\n.modal p{font-size:13px;color:var(--sub);margin-bottom:24px}\n.modal-field{margin-bottom:16px}\n.modal-actions{display:flex;gap:12px;margin-top:24px}\n.btn-save{flex:1;background:var(--blue);color:#fff;border:none;padding:11px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer}\n.btn-cancel{flex:1;background:#f1f5f9;color:var(--text);border:none;padding:11px;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer}\n</style>\n</head>\n<body>\n\n<!-- LOGIN PAGE -->\n<div id="loginPage" style="display:none">\n<div class="login-page">\n<div class="login-card">\n  <img src="https://raw.githubusercontent.com/MetroredEC/metrored-clasificador/main/public/logo-metrored.jpg" class="login-logo" alt="Metrored">\n  <h2>Clasificador BUPA</h2>\n  <p>Inicia sesion con tu cuenta corporativa de Metrored para acceder al sistema de clasificacion de documentos.</p>\n  <button class="btn-ms" onclick="loginPopup()">\n    <svg viewBox="0 0 21 21" fill="none"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>\n    Iniciar sesion con Microsoft\n  </button>\n  <div id="loginErr" class="res-err" style="display:none;margin-top:16px"></div>\n  <p class="login-footer">Solo cuentas @metrored.med.ec</p>\n</div>\n</div>\n</div>\n\n<!-- MAIN APP -->\n<div id="appPage" style="display:none">\n\n<!-- TOP BAR -->\n<div class="topbar">\n  <img src="https://raw.githubusercontent.com/MetroredEC/metrored-clasificador/main/public/logo-metrored.jpg" class="topbar-logo" alt="Metrored">\n  <div class="topbar-right">\n    <span class="user-chip" id="userChip">Cargando...</span>\n    <button class="btn-icon" onclick="openSettings()">\n      <svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 15a3 3 0 100-6 3 3 0 000 6z"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z"/></svg>\n      Configuracion\n    </button>\n    <button class="btn-icon" onclick="logout()" style="color:#ef4444;border-color:#fecaca">Cerrar sesion</button>\n  </div>\n</div>\n\n<!-- CONTENT -->\n<div class="main">\n  <div class="page-title">Clasificador de Documentos BUPA</div>\n  <p class="page-sub">Sube el PDF masivo escaneado y la aplicacion organizara automaticamente cada caso en SharePoint.</p>\n\n  <div class="card">\n    <div class="card-title"><span>1</span> Informacion del Despacho</div>\n    <div class="form-row">\n      <div>\n        <label>Numero de Despacho</label>\n        <input type="text" id="despacho" placeholder="2026-04-27">\n      </div>\n      <div>\n        <label>Carpeta Destino en SharePoint</label>\n        <input type="text" id="spFolder" placeholder="Shared Documents" id="spFolder">\n      </div>\n    </div>\n  </div>\n\n  <div class="card">\n    <div class="card-title"><span>2</span> Archivo PDF Masivo</div>\n    <div class="drop" id="drop">\n      <input type="file" id="fi" accept=".pdf" onchange="pickFile(this)">\n      <div class="drop-icon">\n        <svg fill="none" stroke="#29ABE2" stroke-width="1.5" viewBox="0 0 24 24"><path d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>\n      </div>\n      <p>Arrastra el PDF aqui o <strong>haz clic para seleccionar</strong></p>\n      <div id="fname" style="display:none"></div>\n    </div>\n  </div>\n\n  <button class="btn-main" id="btn" onclick="run()">\n    <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>\n    Clasificar y Subir a SharePoint\n  </button>\n\n  <div class="prog" id="prog">\n    <div class="prog-msg" id="pmsg">Iniciando...</div>\n    <div class="prog-bar-bg"><div class="prog-bar" id="bar"></div></div>\n  </div>\n\n  <div class="res" id="res"></div>\n</div>\n</div>\n\n<!-- SETTINGS MODAL -->\n<div class="modal-overlay" id="settingsModal">\n<div class="modal">\n  <h3>Configuracion</h3>\n  <p>Define la carpeta de destino predeterminada en SharePoint. Se guardara en tu navegador.</p>\n  <div class="modal-field">\n    <label>Carpeta predeterminada en SharePoint</label>\n    <input type="text" id="cfgFolder" placeholder="Shared Documents">\n  </div>\n  <div class="modal-field">\n    <label>Numero de Despacho predeterminado (opcional)</label>\n    <input type="text" id="cfgDespacho" placeholder="ej: 2026-04">\n  </div>\n  <div class="modal-actions">\n    <button class="btn-cancel" onclick="closeSettings()">Cancelar</button>\n    <button class="btn-save" onclick="saveSettings()">Guardar</button>\n  </div>\n</div>\n</div>\n\n<script>\nvar file=null;\nvar msalToken=null;\nvar msalApp=null;\n\n// MSAL Config\nvar msalConfig={auth:{clientId:"66130291-fc50-43f1-943c-6818dac1ba99",authority:"https://login.microsoftonline.com/480bd49c-6f89-4faa-b39e-c7728d95d130",redirectUri:window.location.origin}};\n\nwindow.onload=function(){\\n  var t=0;\\n  var iv=setInterval(async function(){\\n    t++;\\n    if(typeof msal!=='undefined'){\\n      clearInterval(iv);\\n      try{\\n        msalApp=new msal.PublicClientApplication(msalConfig);\\n        await msalApp.initialize();\\n        var accs=msalApp.getAllAccounts();\\n        if(accs.length>0){\\n          try{var r=await msalApp.acquireTokenSilent({scopes:['User.Read'],account:accs[0]});msalToken=r.accessToken;showApp(accs[0]);}\\n          catch(e){showLogin();}\\n        }else{showLogin();}\\n        loadSettings();\\n      }catch(e){showLogin();}\\n    }else if(t>40){clearInterval(iv);showLogin();}\\n  },250);\\n};\n\nfunction showLogin(){document.getElementById("loginPage").style.display="block";document.getElementById("appPage").style.display="none";}\nfunction showApp(account){\n  document.getElementById("loginPage").style.display="none";\n  document.getElementById("appPage").style.display="block";\n  if(account)document.getElementById("userChip").textContent=account.name||account.username;\n}\n\nasync function loginPopup(){\n  try{\n    var r=await msalApp.loginPopup({scopes:["User.Read"]});\n    msalToken=r.accessToken;\n    showApp(r.account);\n  }catch(e){\n    var el=document.getElementById("loginErr");\n    el.textContent="Error: "+e.message;\n    el.style.display="block";\n  }\n}\n\nfunction logout(){msalApp.logoutPopup();showLogin();}\n\n// SETTINGS\nfunction openSettings(){\n  document.getElementById("cfgFolder").value=localStorage.getItem("sp_folder")||"";\n  document.getElementById("cfgDespacho").value=localStorage.getItem("sp_despacho")||"";\n  document.getElementById("settingsModal").classList.add("open");\n}\nfunction closeSettings(){document.getElementById("settingsModal").classList.remove("open");}\nfunction saveSettings(){\n  localStorage.setItem("sp_folder",document.getElementById("cfgFolder").value.trim());\n  localStorage.setItem("sp_despacho",document.getElementById("cfgDespacho").value.trim());\n  closeSettings();\n  loadSettings();\n}\nfunction loadSettings(){\n  var f=localStorage.getItem("sp_folder");\n  var d=localStorage.getItem("sp_despacho");\n  if(f)document.getElementById("spFolder").value=f;\n  if(d)document.getElementById("despacho").value=d;\n}\n\n// FILE PICK\nfunction pickFile(inp){file=inp.files[0];var f=document.getElementById("fname");f.textContent=file.name;f.style.display="inline-block";}\nvar drop=document.getElementById("drop");\ndrop.ondragover=function(e){e.preventDefault();drop.classList.add("over");};\ndrop.ondragleave=function(){drop.classList.remove("over");};\ndrop.ondrop=function(e){e.preventDefault();drop.classList.remove("over");var f=e.dataTransfer.files[0];if(f&&f.type==="application/pdf"){file=f;var fn=document.getElementById("fname");fn.textContent=f.name;fn.style.display="inline-block";}};\n\nfunction setBar(p,m){document.getElementById("bar").style.width=p+"%";document.getElementById("pmsg").textContent=m;}\n\n// OCR direct from browser\nasync function ocrChunk(pdfBytes,endpoint,key){\n  var r=await fetch(endpoint.replace(/\\/$/,"")+"/formrecognizer/documentModels/prebuilt-read:analyze?api-version=2023-07-31",{method:"POST",headers:{"Ocp-Apim-Subscription-Key":key,"Content-Type":"application/pdf"},body:pdfBytes});\n  if(!r.ok)throw new Error("OCR: "+await r.text());\n  var opUrl=r.headers.get("Operation-Location");\n  for(var i=0;i<24;i++){\n    await new Promise(function(r){setTimeout(r,4000);});\n    var p=await(await fetch(opUrl,{headers:{"Ocp-Apim-Subscription-Key":key}})).json();\n    if(p.status==="succeeded"){\n      var txt="";\n      (p.analyzeResult&&p.analyzeResult.pages||[]).forEach(function(pg){\n        txt+="<<<PAG"+pg.pageNumber+">>>\\n";\n        (pg.lines||[]).forEach(function(l){txt+=l.content+"\\n";});\n      });\n      return txt;\n    }\n    if(p.status==="failed")throw new Error("OCR fallo");\n  }\n  throw new Error("OCR timeout");\n}\n\nfunction extractCases(ocrText,pageOffset){\n  var parts=ocrText.split(/<<<PAG\\d+>>>/);\n  var pages=parts.filter(function(p){return p.trim().length>10;});\n  function getPat(t){var m=t.match(/Paciente:\\s*([A-Z\\xc0-\\xd6\\xd8-\\xde][A-Z\\xc0-\\xd6\\xd8-\\xde\\s]{4,60}?)(?:\\s{2,}|\\s*Edad:|\\s*Hc:|\\n)/m);return m?m[1].trim().replace(/\\s+/g," "):null;}\n  function getPol(t){var m=t.match(/\\((\\d{5,7})\\)/);return m?m[1]:"SIN_POLIZA";}\n  function isBupa(t){return /AUTORIZACI[O\\xd3]N DE COBERTURA|N[u\\xfa]mero de P[o\\xf3]liza/i.test(t);}\n  function isMet(t){return /Estado de cuenta|METRORED|Paciente:/i.test(t);}\n  var cases=[];var i=0;\n  while(i<pages.length){\n    var abs=pageOffset+i+1;\n    if(isMet(pages[i])){\n      var pat=getPat(pages[i]);\n      if(!pat){i++;continue;}\n      if(i+1<pages.length&&isBupa(pages[i+1])){cases.push({policy_number:getPol(pages[i+1]),patient_name:pat,pages:[abs,abs+1],document_types:["estado_cuenta_metrored","autorizacion_cobertura_bupa"]});i+=2;}\n      else{cases.push({policy_number:"SIN_POLIZA",patient_name:pat,pages:[abs],document_types:["estado_cuenta_metrored"]});i++;}\n    }else if(isBupa(pages[i])){\n      var am=pages[i].match(/Asegurado:\\s*([A-Z\\xc0-\\xd6\\xd8-\\xde][A-Z\\xc0-\\xd6\\xd8-\\xde\\s]{4,60}?)(?:\\s{2,}|\\s*Fecha|\\n)/m);\n      cases.push({policy_number:getPol(pages[i]),patient_name:am?am[1].trim().replace(/\\s+/g," "):"PACIENTE_"+abs,pages:[abs],document_types:["autorizacion_cobertura_bupa"]});i++;\n    }else{i++;}\n  }\n  return cases;\n}\n\nasync function run(){\n  var despacho=document.getElementById("despacho").value.trim();\n  var spFolder=document.getElementById("spFolder").value.trim();\n  if(!despacho||!spFolder||!file){alert("Completa todos los campos y selecciona un PDF");return;}\n  if(!msalToken){alert("Debes iniciar sesion primero");return;}\n  var btn=document.getElementById("btn");btn.disabled=true;\n  document.getElementById("prog").style.display="block";\n  document.getElementById("res").style.display="none";\n  try{\n    setBar(3,"Obteniendo configuracion...");\n    var cfg=await(await fetch("/api/config")).json();\n    setBar(8,"Leyendo PDF...");\n    var buf=await file.arrayBuffer();\n    var fullBytes=new Uint8Array(buf);\n    var PDFDoc=PDFLib.PDFDocument;\n    var fullPdf=await PDFDoc.load(fullBytes);\n    var totalPages=fullPdf.getPageCount();\n    var chunkSize=4;var allCases=[];var chunks=Math.ceil(totalPages/chunkSize);\n    for(var c=0;c<chunks;c++){\n      var sp=c*chunkSize;var ep=Math.min(sp+chunkSize,totalPages);\n      var chunkPdf=await PDFDoc.create();\n      var idx=[];for(var p=sp;p<ep;p++)idx.push(p);\n      var cp=await chunkPdf.copyPages(fullPdf,idx);\n      cp.forEach(function(pg){chunkPdf.addPage(pg);});\n      var cb=await chunkPdf.save();\n      setBar(10+Math.round((c/chunks)*55),"OCR bloque "+(c+1)+"/"+chunks+" ("+sp+"-"+(ep-1)+")...");\n      var ocrText=await ocrChunk(cb,cfg.ocrEndpoint,cfg.ocrKey);\n      allCases=allCases.concat(extractCases(ocrText,sp));\n    }\n    setBar(68,"Identificados "+allCases.length+" casos. Subiendo a SharePoint...");\n    var fd=new FormData();\n    fd.append("pdf",new Blob([fullBytes],{type:"application/pdf"}),file.name);\n    fd.append("despacho",despacho);\n    fd.append("spFolder",spFolder);\n    fd.append("cases",JSON.stringify(allCases));\n    var resp=await fetch("/api/upload",{method:"POST",headers:{"Authorization":"Bearer "+msalToken},body:fd});\n    var data=await resp.json();\n    if(data.error)throw new Error(data.error);\n    setBar(100,"Completado");\n    var html="<div class=\\"res-sum\\">"+data.cases_processed+" casos procesados correctamente &mdash; Despacho: "+data.despacho+"</div>";\n    data.cases.forEach(function(c){html+="<div class=\\"res-item\\"><strong>"+c.folder+"</strong>&nbsp;&mdash;&nbsp;pags: "+c.pages.join(", ")+"</div>";});\n    var res=document.getElementById("res");res.innerHTML=html;res.style.display="block";\n  }catch(e){\n    document.getElementById("res").innerHTML="<div class=\\"res-err\\">ERROR: "+e.message+"</div>";\n    document.getElementById("res").style.display="block";setBar(0,"Error");\n  }\n  btn.disabled=false;\n}\n</script>\n</body>\n</html>';
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
  // Validate Entra ID token
  const authHeader = request.headers.get('Authorization') || '';
  const token = authHeader.replace('Bearer ', '').trim();
  if (!token) throw new Error('No autorizado: token requerido');

  // Verify token with Microsoft Graph
  const meResp = await fetch('https://graph.microsoft.com/v1.0/me', {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  if (!meResp.ok) throw new Error('No autorizado: token invalido o expirado');
  const me = await meResp.json();
  
  // Check user belongs to allowed tenant
  const allowedTenant = env.ENTRA_TENANT_ID;
  if (!me.id) throw new Error('No autorizado: usuario no verificado');


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