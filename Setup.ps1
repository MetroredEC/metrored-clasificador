# ================================================================
#  CLASIFICADOR METRORED-BUPA  |  Script de Instalacion
#  Ejecutar como Administrador en PowerShell
# ================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function paso  { param($n,$msg) Write-Host "`n=== PASO $n : $msg ===" -ForegroundColor Cyan }
function ok    { param($msg)    Write-Host "  [OK]  $msg" -ForegroundColor Green }
function info  { param($msg)    Write-Host "  [..]  $msg" -ForegroundColor White }
function aviso { param($msg)    Write-Host "  [!!]  $msg" -ForegroundColor Yellow }
function fallo { param($msg)    Write-Host "  [ERROR]  $msg" -ForegroundColor Red; Read-Host "Presiona ENTER para salir"; exit 1 }
function ask   { param($msg)    return (Read-Host "  [?] $msg").Trim() }
function pausa { Read-Host "`n  --> Presiona ENTER para continuar..." | Out-Null }

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   CLASIFICADOR BUPA-METRORED  --  Instalacion              " -ForegroundColor Cyan
Write-Host "   Organiza escaneados masivos en Google Drive con IA       " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
info "Este script te guiara paso a paso. Tiempo estimado: 20 minutos."
pausa

# ================================================================
paso 1 "Verificar Node.js"
# ================================================================
$nodeVer = $null
try { $nodeVer = & node --version 2>$null } catch {}

if (-not $nodeVer) {
    aviso "Node.js no esta instalado."
    info "Descargalo desde: https://nodejs.org/es/download"
    info "Instala la version LTS y cuando termines, vuelve aqui."
    pausa
    try { $nodeVer = & node --version 2>$null } catch {}
    if (-not $nodeVer) { fallo "Node.js sigue sin detectarse. Cierra PowerShell, reinstala Node.js y vuelve." }
}
ok "Node.js detectado: $nodeVer"

# ================================================================
paso 2 "Instalar Wrangler (CLI de Cloudflare)"
# ================================================================
$wranglerVer = $null
try { $wranglerVer = & wrangler --version 2>$null } catch {}

if (-not $wranglerVer) {
    info "Instalando Wrangler globalmente..."
    & npm install -g wrangler
    try { $wranglerVer = & wrangler --version 2>$null } catch {}
    if (-not $wranglerVer) { fallo "No se pudo instalar Wrangler. Intenta: npm install -g wrangler" }
}
ok "Wrangler listo: $wranglerVer"

# ================================================================
paso 3 "Verificar Git"
# ================================================================
$gitVer = $null
try { $gitVer = & git --version 2>$null } catch {}

if (-not $gitVer) {
    aviso "Git no esta instalado."
    info "Descargalo desde: https://git-scm.com/download/win"
    pausa
    try { $gitVer = & git --version 2>$null } catch {}
    if (-not $gitVer) { fallo "Git sigue sin detectarse. Reinstala y vuelve." }
}
ok "Git detectado: $gitVer"

# ================================================================
paso 4 "Clave API de Anthropic"
# ================================================================
Write-Host ""
info "Necesitas tu clave API de Anthropic (la IA que lee los PDFs)."
info "Crearla gratis en: https://console.anthropic.com/settings/keys"
Write-Host ""
$anthropicKey = ask "Pega tu Anthropic API Key (empieza con sk-ant-)"
if ($anthropicKey.Length -lt 10) { fallo "La clave parece muy corta. Intentalo de nuevo." }
ok "Clave de Anthropic guardada"

# ================================================================
paso 5 "Cuenta de Servicio de Google Drive"
# ================================================================
Write-Host ""
Write-Host "  --- Configuracion de Google Drive ---" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Sigue estos pasos en tu navegador:" -ForegroundColor White
Write-Host ""
Write-Host "  PASO A) Crear proyecto en Google Cloud" -ForegroundColor Cyan
Write-Host "     Ir a: https://console.cloud.google.com" -ForegroundColor Gray
Write-Host "     -> Crear proyecto nuevo llamado: metrored-clasificador" -ForegroundColor Gray
Write-Host ""
Write-Host "  PASO B) Activar Google Drive API" -ForegroundColor Cyan
Write-Host "     -> Menu -> APIs y Servicios -> Biblioteca" -ForegroundColor Gray
Write-Host "     -> Buscar 'Google Drive API' -> Habilitar" -ForegroundColor Gray
Write-Host ""
Write-Host "  PASO C) Crear Cuenta de Servicio" -ForegroundColor Cyan
Write-Host "     -> APIs y Servicios -> Credenciales" -ForegroundColor Gray
Write-Host "     -> + Crear Credenciales -> Cuenta de Servicio" -ForegroundColor Gray
Write-Host "     -> Nombre: metrored-service -> Crear y continuar -> Listo" -ForegroundColor Gray
Write-Host ""
Write-Host "  PASO D) Descargar clave JSON" -ForegroundColor Cyan
Write-Host "     -> Clic en la cuenta de servicio creada" -ForegroundColor Gray
Write-Host "     -> Pestana Claves -> Agregar Clave -> Crear nueva clave -> JSON" -ForegroundColor Gray
Write-Host "     -> Se descarga un archivo .json  --- GUARDALO BIEN ---" -ForegroundColor Gray
Write-Host ""
Write-Host "  PASO E) Compartir carpeta de Drive" -ForegroundColor Cyan
Write-Host "     -> Abre Google Drive en el navegador" -ForegroundColor Gray
Write-Host "     -> Clic derecho en la carpeta destino -> Compartir" -ForegroundColor Gray
Write-Host "     -> Pega el email de la cuenta (termina en @...gserviceaccount.com)" -ForegroundColor Gray
Write-Host "     -> Permiso: Editor -> Enviar" -ForegroundColor Gray
Write-Host ""
pausa

$jsonPath = ask "Ruta del archivo JSON descargado (ej: C:\Users\dbermeo\Downloads\metrored-xxxx.json)"
$jsonPath = $jsonPath.Trim('"').Trim("'")

if (-not (Test-Path $jsonPath)) {
    fallo "No se encontro el archivo: $jsonPath"
}

$googleServiceAccount = Get-Content -Raw -Path $jsonPath -Encoding UTF8
try {
    $parsed = $googleServiceAccount | ConvertFrom-Json
    $saEmail = $parsed.client_email
    ok "Cuenta de servicio: $saEmail"
} catch {
    fallo "El archivo JSON no es valido. Asegurate de descargar el archivo correcto de Google Cloud."
}

# ================================================================
paso 6 "Iniciar sesion en Cloudflare"
# ================================================================
Write-Host ""
info "Se abrira el navegador para iniciar sesion en Cloudflare."
info "Si no tienes cuenta, registrate gratis en: https://cloudflare.com"
pausa
& wrangler login
ok "Sesion de Cloudflare iniciada"

# ================================================================
paso 7 "Instalar dependencias del proyecto"
# ================================================================
info "Instalando librerias (pdf-lib)..."
& npm install
ok "Dependencias instaladas"

# ================================================================
paso 8 "Subir codigo a GitHub"
# ================================================================
Write-Host ""
info "Vamos a guardar el codigo en GitHub (respaldo y control de versiones)."
info "Si no tienes cuenta, registrate gratis en: https://github.com"
Write-Host ""

if (-not (Test-Path ".git")) {
    & git init | Out-Null
    ok "Repositorio Git creado"
}

& git add . | Out-Null
& git commit -m "feat: Clasificador BUPA-Metrored" --allow-empty | Out-Null
ok "Codigo registrado en Git"

$ghVer = $null
try { $ghVer = & gh --version 2>$null } catch {}

if ($ghVer) {
    info "Creando repositorio privado en GitHub..."
    try {
        & gh repo create metrored-clasificador --private --source=. --remote=origin --push 2>&1
        ok "Repositorio GitHub creado y codigo subido"
    } catch {
        aviso "No se pudo crear con gh CLI. Continua con el metodo manual."
        $ghVer = $null
    }
}

if (-not $ghVer) {
    Write-Host ""
    aviso "Crea el repositorio manualmente:"
    Write-Host "  1. Ir a: https://github.com/new" -ForegroundColor Gray
    Write-Host "  2. Nombre: metrored-clasificador | Privado | SIN README" -ForegroundColor Gray
    Write-Host "  3. Copia la URL que termina en .git" -ForegroundColor Gray
    Write-Host ""
    $repoUrl = ask "Pega la URL del repositorio (ej: https://github.com/TU_USUARIO/metrored-clasificador.git)"
    & git remote remove origin 2>$null
    & git remote add origin $repoUrl | Out-Null
    & git branch -M main | Out-Null
    & git push -u origin main
    ok "Codigo subido a GitHub"
}

# ================================================================
paso 9 "Guardar claves secretas en Cloudflare"
# ================================================================
Write-Host ""
info "Las claves se guardan encriptadas en Cloudflare, no en el codigo."
Write-Host ""

Write-Host "  Guardando ANTHROPIC_API_KEY..." -ForegroundColor Yellow
Write-Output $anthropicKey | & wrangler secret put ANTHROPIC_API_KEY
ok "ANTHROPIC_API_KEY guardada"

Write-Host ""
Write-Host "  Guardando GOOGLE_SERVICE_ACCOUNT..." -ForegroundColor Yellow
Write-Output $googleServiceAccount | & wrangler secret put GOOGLE_SERVICE_ACCOUNT
ok "GOOGLE_SERVICE_ACCOUNT guardada"

# ================================================================
paso 10 "Desplegar la aplicacion"
# ================================================================
info "Desplegando en Cloudflare Workers..."
$deployLines = & wrangler deploy 2>&1
$deployText  = $deployLines -join " "
$deployLines | ForEach-Object { Write-Host $_ }

$urlMatch = [regex]::Match($deployText, "https://metrored-clasificador\.[a-z0-9\-]+\.workers\.dev")
$appUrl   = if ($urlMatch.Success) { $urlMatch.Value } else { "(busca la URL 'workers.dev' en el texto de arriba)" }

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "   INSTALACION COMPLETADA CON EXITO" -ForegroundColor Green
Write-Host ""
Write-Host "   Tu aplicacion esta en:" -ForegroundColor Green
Write-Host "   $appUrl" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  USO DIARIO:" -ForegroundColor Cyan
Write-Host "  1. Abre la URL de arriba en tu navegador" -ForegroundColor White
Write-Host "  2. Escribe el numero de despacho (ej: 2026-04-08)" -ForegroundColor White
Write-Host "  3. Pega la URL de tu carpeta de Google Drive" -ForegroundColor White
Write-Host "  4. Arrastra el PDF masivo y haz clic en Clasificar" -ForegroundColor White
Write-Host ""
Write-Host "  Para actualizar la app en el futuro, ejecuta:" -ForegroundColor Gray
Write-Host "  npm run deploy" -ForegroundColor Yellow
Write-Host ""

if ($appUrl -notmatch "busca") {
    Set-Clipboard -Value $appUrl
    ok "URL copiada al portapapeles"
}

pausa
