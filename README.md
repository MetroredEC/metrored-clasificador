# 📁 Clasificador BUPA-Metrored

Aplicación web que **analiza escaneados masivos de Metrored/BUPA con IA** y los organiza automáticamente en Google Drive.

## ¿Qué hace?

Toma un PDF masivo con todos los documentos del día y crea esta estructura en tu Drive:

```
📁 DESPACHO-2026-04-08
  📁 700220 - RAMIREZ ANCHUNDIA EDWARD FABRICIO
     📄 1_estado_cuenta_metrored.pdf
     📄 2_autorizacion_cobertura_bupa.pdf
  📁 731473 - MARQUEZ ORDONEZ CAROLINA GISSELA
     📄 1_estado_cuenta_metrored.pdf
     📄 2_autorizacion_cobertura_bupa.pdf
  ...
```

## Instalación (primera vez)

1. Haz clic derecho en `Setup.ps1`
2. Selecciona **"Ejecutar con PowerShell"**
3. Sigue las instrucciones en pantalla (~20 minutos)

## Uso diario

1. Abre la URL de tu aplicación (la obtienes al instalar)
2. Escribe el número de despacho (ej: `2026-04-08`)
3. Pega la URL de la carpeta de Google Drive destino
4. Arrastra el PDF masivo
5. Clic en **"Clasificar y Subir a Drive"**

## Tecnologías

- **Cloudflare Workers** — backend y hosting (gratis hasta 100k requests/día)
- **Claude AI (Anthropic)** — extracción inteligente de datos del PDF
- **Google Drive API** — creación de carpetas y subida de archivos
- **pdf-lib** — separación de páginas individuales

## Límites importantes

| Límite | Plan Gratuito | Plan Pago ($5/mes) |
|--------|--------------|-------------------|
| PDFs pequeños (< 30 páginas) | ✅ | ✅ |
| PDFs grandes (30-200 páginas) | ⚠️ Puede fallar | ✅ |
| Requests por día | 100,000 | 10 millones |

Para PDFs de muchas páginas, se recomienda activar el plan Workers Paid en: https://dash.cloudflare.com

## Actualizar la aplicación

```powershell
# Desde la carpeta del proyecto:
npm run deploy
```

## Solución de problemas

**Error "Google Auth failed"**
→ Verifica que compartiste la carpeta de Drive con el email de la cuenta de servicio

**Error "Claude API error"**
→ Verifica que tu API key de Anthropic sea correcta y tenga saldo

**El PDF no se procesa completo**
→ Divide el PDF en lotes más pequeños o activa el plan Workers Paid

---
*Desarrollado con Claude (Anthropic) para Metrored Centros Médicos*
