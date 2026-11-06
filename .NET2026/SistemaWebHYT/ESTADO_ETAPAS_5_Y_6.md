# Estado de etapas 5 y 6 — SistemaWebHYT

Actualización: 03/11/2026. Implementación y verificación de desarrollo completadas para presupuestos y seguimiento documental. Ver alcance conciliado en VERIFICACION_ETAPAS_5_Y_6.md y uso en GUIA_ETAPAS_5_Y_6.md.

## Etapa 5: presupuestos

Listado, filtros/fechas, ayudas, captura por cliente/tienda, Invitación/Emergencia, partidas manuales/tarifario, PEN/USD, dos decimales, IGV por vigencia y snapshots. Borrador, envío/aprobación y demás transiciones, revisiones históricas y bloqueo de contenido con OC. Emergencia con OC permite regularización comercial sin recalcular importes.

Configuración corporativa HYT y formatos HTML nuevo/anterior conservados por generación con versión, usuario, contenido y SHA256. No se regeneran documentos históricos con maestros actuales.

## Etapa 6: seguimiento

Una OC y una guía por presupuesto, número interno y número OC cliente separados, guía de emergencia sin OC, confirmación desde OC, estados y ejecución con controles de cierre/reapertura. Captura popup, consulta/edición, navegación entre documentos, adjuntos consultables y auditoría paginada. Historial del expediente filtrado por permisos y actualizaciones documentales visibles en presupuestos.

Adjuntos reales en Cloudflare R2, bucket hyt-documentos por configuración; no hay fallback a disco en la aplicación. Turnstile integrado en acceso con validación en servidor y obligatoriedad en producción. Las sustituciones locales y dobles de Cloudflare se usan exclusivamente en pruebas.

## SQL y pruebas

- Fuentes aplicadas directamente a localhost\SQLEXPRESS / SistemaWebHYT.
- 142 scripts, 47 tablas y 105 relaciones validados offline.
- Paquetes incrementales de etapas 5/6: 13 y 20 fuentes, con huellas verificadas.
- 97 comprobaciones SQL reales desde los paquetes en base temporal, eliminada al terminar.
- 164 comprobaciones HTTP/C# aprobadas.
- Chrome verificó presupuesto, seguimiento, adjuntos, aceptación, regularización, filtros y navegación móvil/escritorio.
- Siete casos de formatos HTML/impresión verificados; PDFs de preview solo como evidencia.
- La base de trabajo mantiene cero presupuestos y cero adjuntos de prueba.

## Configuración solicitada por el usuario

User Secrets propio vinculado al proyecto web:

`C:\Users\Franco Lara\AppData\Roaming\Microsoft\UserSecrets\aspnet-SistemaWebHYT-0b5cb513-f035-40a5-b180-58a7e2a649a3\secrets.json`

Contiene ConnectionStrings:HYT para SQL local con autenticación Windows y los campos R2/Turnstile. Conexión SQL verificada. Claves Cloudflare vacías para completar por el usuario; creación del bucket R2 expresamente pendiente a su cargo. La prueba real contra su cuenta Cloudflare se hará después de ese paso. Detalle en CONFIGURACION_CLOUDFLARE.md.

No se creó administrador inicial ni se desplegó el sistema. Aprovisionamiento de usuario según GUIA_ACCESO_ETAPA_2.md; publicación y validación operativa corresponden a etapa 9.

## Siguiente alcance

Etapa 7: marcado Facturado con referencia externa, sin emisión electrónica. Etapa 8: dashboard y reportes. No están incluidos en el cierre 5/6.