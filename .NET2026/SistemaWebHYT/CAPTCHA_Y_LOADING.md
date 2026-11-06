# Captcha de respaldo y espera común HYT

Estado vigente: 05/11/2026. Implementación solicitada por el usuario para mantener el acceso cuando Turnstile no esté disponible y mostrar que las operaciones demoradas siguen en curso.

## Acceso con respaldo manual

- Si falta la clave pública de Turnstile, aparece directamente el captcha manual.
- Con widget configurado se puede elegir «Usar captcha manual». También se activa ante error del script/widget o si no se completa la verificación en doce segundos.
- Si Siteverify rechaza el token o falla la comunicación, el servidor ofrece el captcha manual para un nuevo intento. El fallo no autentica al usuario.
- Cada código tiene seis caracteres, vence en cinco minutos y admite un intento. Se comprueba en el servidor antes de consultar las credenciales. Una contraseña incorrecta sigue impidiendo el ingreso.
- El desafío queda ligado al navegador mediante cookie Secure/HttpOnly; otro navegador no puede obtener su imagen ni validarlo. La solución no se incluye en los campos HTML ni en el JSON de renovación.
- La imagen SVG se entrega sin caché. «Otro código» renueva la imagen conservando el usuario y la contraseña escritos; mientras llega, se indica la espera. Tras un rechazo del servidor, la contraseña se limpia y se genera otro desafío.
- Se conservan antifalsificación, permisos, bloqueo por contraseña incorrecta y auditoría de acceso. Los endpoints de captcha tienen límite de 120 solicitudes/minuto por IP y los accesos mantienen 20/minuto.

ManualCaptchaService mantiene como máximo 5000 desafíos en memoria. El despliegue preparado usa una instancia. Reiniciar el servicio invalida los desafíos pendientes y obliga a pedir otro; para varias réplicas se deberá compartir el almacén. No se añadieron tablas, procedimientos, cuentas o claves ni se cambió la contraseña de ningún usuario en esta ampliación.

## Indicador de espera compartido

Los dos puntos de entrada HTML —login y layout del sistema— cargan hyt-loading.js y hyt-loading.css. El indicador mantiene los colores HYT, aparece tras 150 ms para evitar destellos y avisa si la respuesta supera doce segundos. Se coloca sobre los popups sin robar el foco ni impedir escribir durante cálculos automáticos. Respeta movimiento reducido y no se imprime.

Cubre formularios de acceso/guardado, filtros GET, búsqueda, paginación, actualización de panel/reportes, navegación de la misma pestaña, ayudas, consultas/cálculos AJAX y descargas de adjuntos. Las nuevas pestañas y la ventana de impresión conservan la indicación nativa del navegador; no se inyectan scripts en presupuestos HTML históricos protegidos. Las acciones locales inmediatas, como abrir un calendario, no simulan una espera.

Mantiene un contador de solicitudes simultáneas y espera la lectura del cuerpo de respuesta, no solo sus encabezados. Retira el estado ante éxito, error HTTP, error JSON o cancelación. Los formularios inválidos no quedan ocupados. Los formularios nativos bloquean un segundo envío, conservan el valor del botón enviado y recuperan sus controles al volver con Atrás. Los formularios AJAX conservan sus controles de concurrencia propios; no se confunden con una navegación pendiente. No se permite cerrar un popup mientras su formulario nativo se está enviando.

Las descargas de la pestaña actual esperan la recuperación del archivo antes de entregarlo al navegador, conservando nombre y contenido; un fallo permite reintentar. hytLoading.begin() y hytLoading.navigation() están disponibles para nuevas operaciones; fetch y XMLHttpRequest ya se integran automáticamente.

## Verificación

Compilación aislada sin advertencias ni errores y 205 comprobaciones HTTP/C# aprobadas. Incluyen respuesta correcta/incorrecta, vencimiento, repetición, concurrencia, vínculo al navegador, falta de claves, fallo de Turnstile y contraseña incorrecta después de resolver el captcha.

Chrome verificó escritorio/móvil, imagen legible, renovación, caída del script de Cloudflare, ingreso manual y rechazo de doble envío. Para la espera se simularon solicitudes simultáneas, cuerpo demorado, errores HTTP/JSON, cancelación antes/después de encabezados y XMLHttpRequest, descargas correctas/fallidas, formularios inválidos, popup ocupado, envío nativo/GET y regreso con Atrás. Los datos y archivos de estos recorridos son del host aislado de pruebas; no se conectan a SQL ni a un bucket real.

Capturas en entregables/preview: captcha-manual-escritorio.png, captcha-manual-movil.png, loading-acceso.png y loading-popup.png. La cuenta real del administrador no se usa para estas pruebas.

## Compilar y comprobar

Si Visual Studio está depurando la aplicación, detener primero esa ejecución para recompilar la salida habitual. En esta revisión se utilizó una salida aislada para no interrumpir la sesión abierta:

```text
dotnet build SistemaWebHYT.slnx --artifacts-path entregables/qa-captcha-loading
dotnet entregables/qa-captcha-loading/bin/Hyt.AccessTests/debug/Hyt.AccessTests.dll
dotnet entregables/qa-captcha-loading/bin/Hyt.AccessTests/debug/Hyt.AccessTests.dll --browser-host --browser-captcha
```

Con el último proceso activo, ejecutar `node tools/check_captcha_loading_browser.mjs` en otra terminal. Cerrar el host con Enter. Las pruebas anteriores de presupuestos/documentos usan el mismo host sin `--browser-captcha` y repositorios ficticios; esa variante no altera la configuración del producto.

Para usar el acceso local: detener depuración, recompilar e iniciar GestionServicios.Web con HTTPS. Si las claves siguen pendientes, resolver el captcha manual e ingresar con la cuenta previamente creada. No hace falta desactivar Turnstile ni editar User Secrets para habilitar el respaldo.

Revisión final: recorridos anteriores de presupuestos, seguimiento, mantenimientos y facturado/reportes aprobados con la espera compartida. Cada recorrido que modifica datos ficticios debe comenzar con un host de pruebas nuevo. Release actualizado, activos nuevos incluidos automáticamente por el SDK y 110 archivos inventariados sin secretos. UTF-8/CRLF comprobados en los 22 archivos del cambio; sin commits, cambios SQL ni despliegue.
