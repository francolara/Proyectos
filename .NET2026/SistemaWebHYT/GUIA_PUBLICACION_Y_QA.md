# Preparación de publicación y validación manual — HYT

Fecha: 03/11/2026. Por instrucción expresa del usuario, la publicación queda preparada y no se activa todavía. Faltan claves Cloudflare y pruebas manuales. El dominio sigue parametrizado.

## Estado para el VPS · 10/11/2026

La aplicación y Planillas se han probado localmente. Antes de publicar, quedan estos puntos:

| Punto | Acción pendiente |
|---|---|
| Versión de entrega | Preparar la versión candidata con código web y `Basededatos/Dbsishyt` en una misma revisión. Usar su SHA para identificar la publicación, el paquete SQL y la imagen. |
| Imagen Docker | La matriz de `build-images.yml` incluye HYT como `ghcr.io/francolara/hyt`. Cuando se publique la revisión, fijar `HYT_IMAGE` a la etiqueta del SHA y verificar acceso desde el VPS. La activación del servicio HYT sigue el paso manual separado de Contabo. |
| SQL del servidor | Confirmar la versión del motor y si `SistemaWebHYT` ya existe. El manifiesto requiere SQL Server 2022 como mínimo. Elegir instalación limpia o actualización de una copia restaurada antes de tocar producción. |
| Permisos SQL | El manifiesto incorpora `047_PermisosEntregaPlanillas.sql`. Verificar en el VPS con el usuario real de la aplicación, con rol `hyt_app_executor`. |
| Infraestructura | Confirmar dominio, puerto libre, redes Docker, IP del proxy, certificado y DNS proxied. Completar R2, Turnstile y el usuario SQL en el archivo privado del servidor. |
| Respaldo y aceptación | Verificar que el trabajo de respaldo incluya HYT. Probar restauración de SQL, adjuntos R2 y claves persistentes. Ejecutar la aceptación de Planillas antes de habilitar el dominio. |

### Base de datos y artefactos de Planillas

El inventario de SQL vigente es `Basededatos/Dbsishyt/manifest.json` desde la raíz del repositorio Proyectos. Para una base nueva, ensayar el manifiesto completo en su orden. Para una base existente, inventariar el esquema y los paquetes aplicados y preparar el incremento exacto sobre una copia restaurada. Los paquetes históricos de Planillas y `SHA256_PUBLICACION.json` son referencias de entregas anteriores. La entrega nueva debe cubrir reingresos, tasas previsionales, confirmación de planilla y cobertura opcional de pagos a proveedores. Verificar que el paquete SQL y la imagen pertenezcan a la misma revisión.

El script `047_PermisosEntregaPlanillas.sql` concede ejecución de reingreso, consulta de estimados, edición de tasas y confirmación de pagos de planilla al rol `hyt_app_executor`. La prueba aislada verifica los cuatro permisos efectivos y que el rol no puede actualizar directamente la tabla de trabajadores.

El validador estático reconoce las adiciones de varias columnas en una misma sentencia. Ejecutar `Hyt.SqlValidation`, `Hyt.DatabaseTests --run-local-sql` y `Hyt.AccessTests` antes de regenerar los artefactos. La comprobación SQL se hace en una base de prueba aislada. En producción, repetir las verificaciones con el usuario restringido y los permisos funcionales asignados.

La carpeta de diagramas de Contabo describe la ruta de FRALSE desde GitHub Actions y GHCR hasta despliegue manual, Nginx detrás de Cloudflare, SQL en red privada y respaldos cifrados a R2. HYT conserva su Compose propio. El workflow de despliegue existente todavía no tiene HYT como destino. Su incorporación requiere coordinación con el comando autorizado del VPS.

### Aceptación adicional de Planillas

- Registrar un trabajo desde Presupuestos y confirmar un abono parcial y uno total desde **Pagos a Proveedores**. Verificar el saldo y el historial.
- Registrar y confirmar una obligación AFP u ONP. Verificar el cálculo del estimado, sus trabajadores y el número de operación.
- Confirmar una planificación de prueba. Comprobar los movimientos del personal, el bloqueo de edición y que el reintento conserve un solo pago.
- Revisar cese y reingreso, cuentas bancarias según permisos y SCTR. Comparar vista previa e impresión horizontal de los tres reportes.
- Verificar que un perfil de solo consulta pueda abrir los listados autorizados y que sus acciones de registro estén bloqueadas.

Esta revisión se hizo sobre el repositorio y SQL local. La disponibilidad de redes, dominio, secretos y respaldos en Contabo queda por verificar en el VPS.

## Archivos de entrega

- Dockerfile en la raíz del proyecto: compilación .NET 10 y ejecución ASP.NET Core 10 sin usuario root.
- .dockerignore: excluye secretos, certificados privados, herramientas, pruebas y entregables del contexto de imagen.
- deploy/compose.hyt.yml: servicio HYT independiente, puerto solo en loopback, redes externas SQL/salida y volumen de claves persistentes.
- deploy/.env.compose.example: dominio, imagen versionada, puerto, IP de proxy y nombres de redes que se deben completar/verificar.
- deploy/hyt.env.example: plantilla de configuración privada SQL/R2/Turnstile.
- deploy/nginx.hyt.conf.example: servidor HTTPS parametrizado para revisión, sin activar.
- entregables/publicacion: salida portable de dotnet publish Release, sin secretos. La imagen Docker se construye desde las fuentes.

El proyecto usa la infraestructura de DOCUMENTACION_INFRAESTRUCTURA_FRALSE.md. Las redes se verifican en el servidor; no se asume que sus nombres reales sean backend/egress. El puerto propuesto 5186 debe comprobarse libre antes de adoptarlo. La plantilla utiliza un dominio reservado .invalid y una IP de proxy sin completar para evitar confundir valores de ejemplo con configuración operativa.

## Requisitos antes de activar

1. Confirmar dominio y preparar su DNS/proxy de Cloudflare y certificado Origin CA, conservando la protección del origen ya definida en FRALSE.
2. Crear el bucket privado hyt-documentos y completar credenciales R2 con lectura/escritura de objetos limitada al bucket.
3. Configurar widget Turnstile, sus claves y dominios permitidos. En producción es obligatoria la verificación de acceso, con Turnstile y su captcha manual de respaldo.
4. Preparar la base SistemaWebHYT en SQL del servidor y aplicar los paquetes/fuentes en orden. Crear un login SQL de aplicación independiente y su usuario de base; asignar hyt_app_executor. No utilizar la autenticación Windows de desarrollo dentro del contenedor Linux ni credenciales administrativas en la aplicación.
5. Aprovisionar primer administrador según GUIA_ACCESO_ETAPA_2.md y la herramienta Hyt.Admin. Las contraseñas se suministran por configuración privada, nunca en scripts/versionado ni argumentos de publicación.
6. Copiar hyt.env.example a un archivo privado de servidor, por ejemplo /srv/fralse/secrets/hyt.env; completar claves/conexión y restringir acceso. ASP.NET Core recibe esas variables mediante env_file. User Secrets solo corresponde a desarrollo local.
7. Verificar nombres de redes, IP del proxy vista por el contenedor y puerto. Configurar ReverseProxy:KnownProxy únicamente con esa IP; el arranque de producción rechaza una IP vacía o inválida. No configurar confianza en todas las redes.
8. Preparar una imagen versionada y conservar la versión anterior para reversión.

## Preparación local reproducible

Desde la raíz SistemaWebHYT:

```text
dotnet publish GestionServicios.Web/GestionServicios.Web.csproj -c Release -o entregables/publicacion /p:UseAppHost=false
docker compose --env-file deploy/.env.compose.example -f deploy/compose.hyt.yml config --no-env-resolution -q
docker build -t hyt-web:VERSION_VERIFICADA .
```

La última instrucción requiere motor Docker disponible. No se ejecutó esa construcción en esta entrega porque el motor local no estaba disponible. No se inició Docker Desktop ni un contenedor. La validación Compose no confirma que redes/bucket/dominio existan: solo comprueba la configuración.

## Activación posterior, únicamente tras pruebas y autorización

Ubicaciones propuestas: configuración Compose de HYT en /srv/fralse/compose/compose.hyt.yml; entorno privado en /srv/fralse/secrets/hyt.env; parámetros no secretos en /srv/fralse/compose/hyt.compose.env. Ajustar rutas según la organización final.

1. Respaldar base, objetos/documentos y claves persistentes; guardar identificación de versión y huellas de paquetes.
2. Revisar y completar las plantillas. Verificar que la aplicación tenga acceso de escritura al volumen de Data Protection como UID del usuario app de la imagen.
3. Construir o descargar la imagen versionada y verificarla en ambiente de pruebas, con claves y base propias.
4. Validar Compose con sus parámetros reales sin imprimir valores secretos en reportes compartidos. Revisar nginx -t antes de habilitar el host.
5. Iniciar únicamente el servicio HYT cuando el usuario autorice la activación. Conservar redes/servicios existentes; no eliminar volúmenes.
6. Verificar HTTPS, esquema de proxy, cookies seguras, login/Turnstile y recorridos de negocio.

La sonda GET /health/live responde ok sin información de conexión ni datos privados; indica que el proceso está vivo. No sustituye la verificación de SQL, R2 o Turnstile. El proxy reenvía protocolo/IP de una sola etapa y la aplicación solo acepta esos encabezados desde su IP verificada.

## Pruebas manuales de aceptación

- Acceso con Turnstile válido/inválido, cuenta desactivada y sesión cerrada; sin claves no omitir el desafío.
- Presupuesto Invitación: cliente/tienda, tarifa/manual, PEN y USD pendiente, IGV, envío/aprobación y dos formatos HTML.
- Emergencia: guía sin OC, ejecución sin programación, OC posterior y regularización comercial sin alterar importes.
- OC: número interno/cliente separados y adjunto privado, confirmación con fecha/número/archivo, consulta posterior y permisos de solo lectura.
- Guía: aceptar con representante/fecha, observar/anular y comprobar dependencias al cerrar/reabrir OC/trabajo.
- Facturado: bloquear sin requisitos, registrar 01/03 con ceros, corregir y retirar con motivo; verificar usuario, fechas y datos anteriores en auditoría.
- Dashboard: mismo rango de fechas que listados, conteos conciliados, PEN/USD separados, series de curvas y detalle de valores. Consultar e imprimir los cuatro reportes.
- Adjuntos: cargar/visualizar/descargar en R2, retiro auditado, tipo/tamaño inválidos, acceso sin permiso y recuperación tras reiniciar la aplicación.
- Dos sesiones: editar el mismo registro y verificar rechazo de la versión obsoleta; doble clic/reenvío sin duplicados.
- Móvil: filtros, ayudas, calendarios, tablas, popups, gráficos y retorno de foco.
- Reinicio del servicio: claves de sesión persistidas y recursos privados recuperables.

Registrar resultados, usuario responsable, fecha, versión probada y evidencias sin credenciales ni datos personales innecesarios.

## Respaldo y restauración

Antes de publicar, incorporar SistemaWebHYT al esquema de backup SQL existente documentado para FRALSE (03:00 Perú, verificación y cifrado). Verificar inclusión real y retención; no asumir cobertura por el nombre de la base.

El respaldo coordinado debe cubrir SQL (referencias/metadatos/historia), objetos R2 y volumen de Data Protection. Un .bak no contiene los archivos adjuntos. R2 no debe tratarse como sustituto de un respaldo independiente: preparar exportación/copia con inventario de claves y SHA256 a almacenamiento privado de respaldo, con credenciales separadas. Proteger/cifrar los respaldos de claves y configuración.

Ensayar restauración en una base y almacenamiento de prueba: restaurar backup con verificación/checksum, recuperar objetos, comprobar recuentos/vínculos/huellas, abrir presupuestos históricos y adjuntos y ejecutar recorridos. No reutilizar permisos de escritura sobre el bucket productivo durante un ensayo de recuperación. Documentar tiempo y resultado.

## Reversión

Conservar imagen anterior, copia de parámetros privados y respaldo previo. Si falla la nueva imagen, detener únicamente HYT y volver a la imagen verificada manteniendo el volumen de claves. Revisar compatibilidad de esquema antes de arrancar una versión anterior; una reversión de imagen no revierte SQL. Restaurar base/objetos solo mediante procedimiento revisado, con copia de seguridad del estado actual y autorización expresa, evitando perder registros creados después del respaldo.

## Verificado en esta preparación

Compilación Release, ausencia de archivos secretos en publicación y validación Compose aprobadas. Pruebas de sonda de vida y configuración de proxy aprobadas: producción exige IP verificada y no confía en redes completas. Pendientes reales: construir/ejecutar imagen con motor Docker disponible, claves/bucket/dominio, aceptación manual, ensayo de restauración y activación autorizada. No se declara publicado ni validado contra infraestructura productiva.

Referencias: [contenedores ASP.NET Core 10](https://learn.microsoft.com/en-us/aspnet/core/host-and-deploy/docker/building-net-docker-images?view=aspnetcore-10.0), [usuario no root e imágenes .NET](https://learn.microsoft.com/en-us/dotnet/core/docker/).

## Consolidación de entrega — 05/11/2026

VERIFICACION_ETAPAS_7_8_Y_9.md registra cobertura y pruebas reproducibles: 136 comprobaciones SQL, 186 HTTP/C#, navegador y cuatro reportes impresos. Los listados de presupuestos/seguimiento y reportes conservan cliente/tienda/estado al navegar desde el panel. Verificar manualmente con documentos reales de HYT y permisos de cada perfil cuando el ambiente esté habilitado. Capturas y PDF actuales contienen únicamente datos ficticios de QA.

Revalidación Docker del 05/11: CLI no disponible y comprobación adicional con validador temporal rechazada automáticamente («blocked by policy»). Revisión estática realizada; ejecutar Compose config y construir/probar imagen en el ambiente habilitado antes de activar. No se afirma ejecución nueva de esos comandos. Inventario de salida Release: entregables/SHA256_PUBLICACION.json.

## Verificación de respaldo y espera — 05/11/2026

La verificación obligatoria de acceso admite Turnstile o captcha manual de respaldo por indicación del usuario. Validar además: claves ausentes, script bloqueado, código vencido/incorrecto, contraseña incorrecta con captcha válido y recuperación de un código nuevo. Revisar la espera en búsquedas, guardados, descargas, popups, error y retorno con Atrás. CAPTCHA_Y_LOADING.md documenta las pruebas automatizadas; el respaldo no reemplaza configurar las claves reales del proveedor.
