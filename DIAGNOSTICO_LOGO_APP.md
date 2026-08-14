# Diagnóstico: App "Seguimiento de tareas" en blanco / logo roto

## Contexto del proyecto
- Repo: `normabersoza-spec/Evaluacion-de-proveedores`
- Rama de trabajo: `claude/trabajo-con-codigos-ey5gon`
- App: Google Apps Script Web App ("Seguimiento de tareas" / Team Center de sourcing de Criotec), servida desde `Code.gs` (`doGet()` → `renderDashboardSourcing()` y `renderVistaSolicitante()`), con vistas `Dashboard.html`, `Dashboard-Escritorio.html`, `Dashboard-Movil.html`, `Solicitudes.html`.

## Problema original
La app tardaba ~8.4s en cargar (438 kB) y a veces se quedaba en blanco.

## Causa raíz #1 (resuelta): logo embebido como base64 duplicado
`doGet()` llamaba a `getLogoDataUri_()`, que descarga el logo (~2MB PNG) de Drive y lo mete como texto base64 dentro del HTML. Como el logo aparece dos veces (vista escritorio + vista móvil), esto duplicaba ~2MB en cada carga. Ya existía una función `getLogoUrl_()` (URL de Drive en vez de base64) creada previamente para este mismo problema, pero nunca se conectó a `doGet()`.

**Fix aplicado:** cambiar las dos llamadas en `doGet()` (`renderDashboardSourcing()` y `renderVistaSolicitante()`) de `getLogoDataUri_()` a `getLogoUrl_()`. Commit `ea3b5b3`.

## Causa raíz #2 (resuelta en código, pendiente de acción manual): escalado de permisos de Drive
La primera versión de `getLogoUrl_()` incluía `logoFile.setSharing(...)` para auto-compartir el archivo si estaba privado. Apps Script calcula los scopes de OAuth de **todo el proyecto de forma estática** (analiza el código, no solo lo que se ejecuta en runtime). Tener esa línea en el código obliga a pedir el scope de **escritura** de Drive (`drive`) en vez de solo lectura (`drive.readonly`). Como ese scope ampliado nunca se autorizó para la implementación publicada, `doGet()` fallaba y Apps Script devolvía una página en blanco (solo `<meta http-equiv="X-UA-Compatible">`, sin error visible) — el fallo típico de Apps Script cuando falta autorización.

**Fix aplicado:** se quitó `setSharing()` de `getLogoUrl_()`; ahora solo arma la URL del thumbnail (`https://drive.google.com/thumbnail?id=...&sz=w400`), sin pedir permisos de escritura. Commit `5c7685d`.

```js
// Code.gs (ya en la rama)
function getLogoUrl_() {
  return "https://drive.google.com/thumbnail?id=" + LOGO_FILE_ID + "&sz=w400";
}
```

## Pendiente / próximos pasos
1. **Acción manual del usuario (aún no confirmada):** compartir el archivo del logo en Drive como "Cualquier persona con el enlace" (rol Lector).
   - Archivo: `LOGO_FILE_ID = '17EUZTytjrow3KR_EzkEYeaA2cUCEyMhk'`
   - URL: `https://drive.google.com/file/d/17EUZTytjrow3KR_EzkEYeaA2cUCEyMhk/view`
   - Sin este paso, el logo seguirá roto porque el thumbnail de Drive no carga si el archivo es privado.
2. **Volver a implementar** el script en Apps Script (pegar `Code.gs` actualizado → Implementar > Gestionar implementaciones > Nueva versión).
3. **Verificar** que la app cargue y el logo se vea. Si la página sigue en blanco después de estos dos pasos, el problema ya no sería el logo — pedir al usuario abrir DevTools (F12) → pestaña Console y compartir cualquier error en rojo.

## Alternativa pendiente (Opción B, no completada)
Se propuso embeber el logo como base64 comprimido **una sola vez** (compartido por JS entre las dos vistas) para eliminar por completo la dependencia de Drive. No se pudo completar en la sesión anterior porque:
- El entorno de esa sesión no tenía salida de red hacia `drive.google.com` (bloqueado por política de proxy).
- La imagen que el usuario pegó en el chat no llegó como archivo accesible al sistema de archivos de esa sesión (solo se pudo "ver", no procesar).

Si se retoma esta opción en la nueva conversación: pedir al usuario que **adjunte el archivo de imagen como archivo** (no solo pegado en el chat) para poder leerlo y comprimirlo, o que confirme si el nuevo entorno sí tiene salida de red hacia Drive para poder descargarlo directamente con el `LOGO_FILE_ID` de arriba.

## Notas del catálogo de funciones (por si es útil)
- `getLogoDataUri_()` (línea ~25 de `Code.gs`): versión vieja, base64, se deja disponible solo para usos como correos (donde no se puede usar URL externa). No usar en `doGet()`.
- `testLogoActual()`: función de prueba que llama a `getLogoDataUri_()`, no afecta la app web.
