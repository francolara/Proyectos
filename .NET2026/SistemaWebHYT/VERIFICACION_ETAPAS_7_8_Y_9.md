# Verificación de etapas 7, 8 y preparación de 9

Estado vigente: 05/11/2026. Las notas anteriores de avance son históricas. El usuario indicó dejar preparada la publicación porque faltan claves y pruebas manuales; no se ha activado el servicio.

| Requisito | Implementación y evidencia |
|---|---|
| Marcado externo Facturado | InvoiceMarkController, IInvoiceMarkRepository, spu_Facturado_Guardar/Obtener. Fecha, códigos textuales 01/03, serie y número alfanuméricos; conserva ceros. No emite comprobantes. |
| Requisitos y seguridad | Requiere OC válida, confirmación vigente, ejecución terminada y guía aceptada; permisos en MVC y SQL. Antifalsificación, versiones y GUID de operación. Pruebas SQL de rechazo, reintento y corrección. |
| Corrección/retiro | Motivo obligatorio y conservación del historial; auditoría de anterior/nuevo, usuario y fechas. Pruebas SQL y recorrido Chrome de alta, corrección y retiro. |
| Visibilidad | Referencia vigente en listado OC solo con permiso específico; ficha enlazada e historial paginado de veinte. Prueba SQL de ocultación sin FACTURADO.CONSULTAR. |
| Dashboard conectado | HomeController y spu_Dashboard_Obtener. Última revisión por presupuesto y una fila por OC; pruebas SQL de conteos y monedas. Fuentes limitadas por permisos. |
| Filtros y navegación | Último mes móvil, fechas Perú, cliente/tienda por nombre histórico. SQL aplica filtros antes de conteos y totales. Se conservan en listados/reportes y al paginar; columnas por estado abren el listado filtrado. Pruebas SQL y Chrome. |
| Curvas y movimientos | SVG con selección de series, detalle por teclado/tacto y día/semana/mes. Hasta noventa intervalos, identificados cuando agrupan varios días o meses. Primer registro de confirmación/marcado por OC, independiente de la fecha de registro de la OC; correcciones/retiros no duplican ese evento. Pruebas SQL de primeras incorporaciones y Chrome de agrupaciones sin pérdida de cantidades. |
| Columnas, dona y vistas rápidas | Estados por fuente autorizada; dona con categorías disjuntas de OC activas. Indicadores de sin confirmar, confirmadas sin marcado, marcadas y listas para marcar. Accesos conservan fechas/filtros; tablas alternativas accesibles. |
| Pendientes históricos | Presupuestos sin OC/guía y emergencias por regularizar, antigüedad y enlace al expediente. Muestra conteos del período y acumulados hasta Hasta; página SQL de veinte, orden antiguo primero. Exige consulta de presupuestos, OC y guías. Pruebas de más de veinte filas y restricción por permiso. |
| Reportes HTML | GANADOS, OC_COMPLETADAS, OC_CON_ACTA y OC_PENDIENTES_FACTURA. Veinte filas SQL; conteo e importes de toda la consulta separados PEN/USD. Categorías superpuestas identificadas. Pruebas SQL de paginación, totales y exclusión de facturadas. |
| Impresión y móvil | Cuatro reportes impresos desde HTML en Chrome, A4 horizontal, veinte filas completas en dos páginas físicas por caso y encabezados repetidos. Totales globales identificados; botón imprime solo la página consultada del sistema. Capturas escritorio/móvil revisadas; tablas móviles con desplazamiento horizontal. |
| Vacío/error/permisos | Resultados SQL vacíos devuelven arreglos vacíos, sin omitir propiedades del panel. Estados sin fuentes, sin registros y sin series probados en Chrome; perfiles sin reportes no consultan el repositorio. Fechas/páginas/longitudes inválidas rechazadas. |
| Preparación de entrega | Dockerfile, Compose, plantillas privadas, Nginx, proxy explícito, sonda de vida y guía GUIA_PUBLICACION_Y_QA.md. Publicación Release sin secretos; revisión estática de configuración y comando Compose preparado para el ambiente habilitado. Sin despliegue remoto ni cambios a servicios existentes. |

## Resultado de validación

- 136 comprobaciones SQL aprobadas instalando los paquetes incrementales en una base temporal con nombre GUID, eliminada al terminar. No se añadieron registros comerciales ficticios a SistemaWebHYT.
- 186 comprobaciones HTTP/C# aprobadas, con repositorios aislados y simulaciones para Cloudflare.
- Validación offline: 148 fuentes, 47 tablas, 105 relaciones; cinco regresiones negativas del validador detectadas.
- 82 procedimientos de la base local comparados con sus fuentes. Solo se normalizan espacios y CREATE OR ALTER frente a CREATE al comparar el texto que almacena SQL Server.
- Paquetes ETAPA_5/6/7/8_HYT.sql y sus inventarios SHA256 verificados. Las fuentes individuales siguen siendo la autoridad.
- Chrome: cuatro series, agrupaciones, navegación por estado/cliente/tienda, datos vacíos/parciales, reportes, escritorio/móvil y corrección/retiro Facturado. Los datos de capturas y PDF son ficticios del host aislado.
- PDFs de los cuatro reportes: todos los renglones una vez, totales, encabezados y texto dentro del papel. Los presupuestos mantienen sus formatos anterior/nuevo y sus verificaciones previas independientes.

## Reproducción local

Desde la raíz SistemaWebHYT:

```text
dotnet build SistemaWebHYT.slnx
dotnet run --project tools/Hyt.DatabaseTests -- --run-local-sql --use-bundles
dotnet run --project tools/Hyt.AccessTests
dotnet run --project tools/Hyt.SqlValidation -- ../../Basededatos/Dbsishyt
python tools/test_sql_validation.py
python tools/check_installed_modules.py
python tools/package_documents.py
```

La revisión visual usa `dotnet run --project tools/Hyt.AccessTests -- --browser-host` en una terminal interactiva local, seguido de `node tools/check_dashboard_browser.mjs` y `node tools/check_delivery_browser.mjs`. Cerrar con Enter; se retiran las credenciales ficticias temporales. `python tools/check_delivery_pdfs.py` verifica los cuatro PDF generados. Estas herramientas no usan una cuenta Cloudflare real ni datos comerciales.

## Pendiente antes de activar

Bucket privado hyt-documentos, claves R2/Turnstile, dominio y parámetros reales de servidor; construir y ejecutar la imagen con Docker disponible; primer administrador; aceptación manual y ensayo de restauración de SQL, objetos y claves. La guía de publicación enumera los pasos y la reversión. La configuración preparada no demuestra conectividad con producción ni una restauración real.

No se han incorporado claves, credenciales por defecto, un bucket o un dominio operativo. La conexión de desarrollo permanece en el User Secrets propio del proyecto, fuera del repositorio.

## Límite de validación Docker — 05/11/2026

En esta sesión el ejecutable Docker no está disponible. La revisión automática rechazó una comprobación adicional que descargaba/ejecutaba un validador temporal y lo eliminaba, con el único motivo «blocked by policy». Se realizó revisión estática de las definiciones; no sustituye ejecutar Docker Compose config, construir o probar la imagen en el ambiente habilitado. Los avances del 03/11 conservan su validación histórica, pero no se presentan como una nueva ejecución del 05/11.

La salida Release contiene 104 archivos inventariados en entregables/SHA256_PUBLICACION.json; no contiene secrets.json, archivos .env ni certificados privados. Se comprobó UTF-8/CRLF en 68 archivos cambiados hoy; AGENTS.md conserva su codificación original. No se realizaron commits.
