
## Etapa 7 iniciada — 03/11/2026

Aplicados directamente en SistemaWebHYT local spu_Facturado_Guardar, spu_Facturado_Obtener y Script/018_PermisosEtapa7.sql. No se crearon tablas nuevas: se utiliza ReferenciaComprobanteExterno y DocumentoOperacion. Generador tools/prepare_invoice_marks.py; manifiesto incremental etapa7 con tres fuentes; acumulado 145 fuentes.

Guardar registra fecha, código textual 01/03, serie y número alfanuméricos, vinculados a OC y confirmación vigente. Mantiene los requisitos documentales existentes del trigger 007: OC válida, ejecución terminada y guía aceptada. No emite comprobantes ni asientos ni se comunica con SUNAT. Corrección y retiro requieren motivo y versiones; retiro conserva referencia anterior. Auditoría y actualización de OC/presupuesto son atómicas. Operación GUID con huella evita duplicados y rechaza reutilización con otros datos. Obtener devuelve contexto, referencia vigente e historial de veinte referencias por página.

Validación: 102 comprobaciones SQL reales aprobadas en base temporal eliminada. Cinco nuevos casos prueban bloqueo sin confirmación, código 01 y ceros conservados, reintento idempotente, una referencia vigente y retiro con motivo/auditoría. Fuente corregida y aplicada para validar requisitos antes de que actúe el trigger de protección existente. Sin datos comerciales de prueba en la base de trabajo.

Pendiente del objetivo 7/8/9: repositorio y popup/listado de marcado externo, auditoría visible y pruebas web; dashboard y reportes HTML conciliados por moneda; validación integral y preparación de entrega/despliegue. Bucket R2 y claves Cloudflare siguen pendientes del usuario conforme a sus instrucciones. Objetivo activo, sin declaración de cierre.
## Integración MVC del marcado externo — 03/11/2026

IInvoiceMarkRepository y SqlInvoiceMarkRepository consumen exclusivamente spu_Facturado_Obtener/Guardar. InvoiceMarkController expone /OCCliente/{id}/Facturado con permisos FACTURADO y OC_CLIENTE, antifalsificación, tokens de concurrencia y GUID de operación. El listado OC ofrece botón popup Facturado; reutiliza calendario, ventana y control de doble envío existentes. Guarda fecha, código 01/03, serie y número; permite corrección y retiro con motivo. Muestra historial de referencias vigente/retiradas, con página SQL de veinte. No emite comprobantes ni los envía a SUNAT.

Pruebas HTTP nuevas aprobadas: parcial popup, campos/códigos, conservación de ceros, retorno local seguro, rechazo de registro incompleto y consulta sin edición. Pruebas SQL ampliadas con repositorio real para lectura de referencia retirada y marcado Boleta código 03. Compilación sin errores ni advertencias. Pendientes de etapa 7: exponer referencia y estado en grilla/ficha OC, auditoría detallada visible, pruebas de retiro/corrección en navegador y conciliación final. Etapas 8 y 9 continúan pendientes; objetivo activo.
## Visibilidad de referencia y auditoría — 03/11/2026

spu_Facturado_Obtener incorpora auditoría detallada y total de eventos, en páginas de veinte, con usuario, fechas, motivo y JSON anterior/nuevo. El formulario MVC muestra las acciones y los cambios conservados. La ficha OC enlaza directamente a Facturado.

Los SP Seguimiento_Listar incorporan un objeto Facturado con fecha, tipo, serie, número, fecha/usuario de marcado. El objeto se devuelve únicamente al tener FACTURADO.CONSULTAR; otros permisos de OC no revelan estos datos. Grilla MVC muestra columnas adicionales únicamente con ese permiso. Fuentes modificadas aplicadas directamente en SQL local, conservando manifiestos de etapas posteriores al regenerar seguimiento.

Verificación: 107 comprobaciones SQL reales aprobadas, incluidas corrección con valor anterior consultable, referencia actual en listado y ocultación ante falta de permiso específico. Pruebas MVC aprobadas y Chrome verificó captura/guardado desde popup OC, recorrido documental previo y escritorio/móvil. Host temporal cerrado. Compilación sin advertencias/errores. Pendiente: ampliar casos de retiro/corrección de UI y consolidar paquetes de etapa7; comenzar dashboard/reportes etapa8 y preparación/validación etapa9. Objetivo activo.
## Etapa 8: consultas de indicadores y reportes — 03/11/2026

Aplicados directamente spu_Dashboard_Obtener, spu_Reporte_Listar y Script/019_PermisosEtapa8.sql. Generador prepare_reporting.py y manifiesto etapa8 con tres fuentes. No se crearon tablas; acumulado 148 fuentes, 47 tablas y 105 relaciones. Generadores de etapas anteriores conservan los manifiestos posteriores al regenerarse.

Dashboard usa cohorte por fecha de registro de presupuesto/OC, con rango UTC enviado desde fechas Perú. Cada presupuesto cuenta una sola vez con su última revisión; cada OC una vez con su revisión vinculada. Importes separados PEN/USD; presupuestos sin precio se contabilizan explícitamente. Indicadores: estados, sin confirmar, confirmadas sin marcado, listas para marcar (confirmación + ejecución terminada + guía aceptada), facturadas y movimientos diarios agrupados en fecha Perú. Requiere REPORTES.CONSULTAR y limita cada fuente por permiso de presupuestos/OC. No mezcla monedas ni suma categorías superpuestas.

Reporte_Listar ofrece GANADOS (última revisión aprobada por presupuesto, excluido expediente cancelado), OC_COMPLETADAS (ejecución terminada), OC_CON_ACTA (guía no anulada) y OC_PENDIENTES_FACTURA (OC no cancelada sin referencia vigente). Categorías superpuestas, no sumables. Cohorte por registro, veinte filas SQL por página, conteo y totales de toda la consulta por moneda. Pendientes distinguen sin confirmar, trabajo pendiente, guía pendiente y lista para marcar.

Pruebas reales: 117 comprobaciones aprobadas en base temporal eliminada; nuevas verifican conteo sin duplicar revisiones/confirmaciones históricas, separación PEN/USD, paginación sin duplicados de cuatro reportes, exclusión de OC facturada, permisos y cambio de día UTC/Perú. Validación offline aprobada. Pendientes: repositorios y dashboard MVC con gráficos profesionales/interacción, reportes HTML y conciliación visual; cierre de pruebas/paquetes etapa7; preparación y validación etapa9. Objetivo activo.
## Reportes MVC y decisión de publicación — 03/11/2026

Implementados IReportingRepository/SqlReportingRepository por ADO.NET con SP; ReportingDates convierte límites inclusivos/exclusivos desde fechas Perú a UTC. ReportsController y /Reportes ofrecen catálogo cerrado de cuatro reportes, permisos REPORTES y documento fuente, rango inicial último mes, veinte filas y totales de toda la consulta por moneda. HTML codifica datos y enlaza presupuesto/OC; imprime la página consultada en A4 horizontal con encabezado de tabla repetido. Menú Reportes redirige a esta vista. No hay consultas SQL en Razor.

Pruebas HTTP aprobadas para cuatro vistas, paginación, ceros/códigos, importes, codificación de datos y permisos. Pruebas con repositorios SQL reales incorporadas para filas/conteos/totales y dashboard autorizado. Pendientes de etapa8: dashboard visual del inicio con curvas, barras/dona, interacciones, vistas rápidas y revisión visual/impresión de reportes. Etapa7 conserva pendientes de cierre de casos UI/paquetes.

Instrucción expresa del usuario para etapa9: dejar la publicación preparada; no publicar todavía porque faltan claves y pruebas manuales. No solicitar nuevamente dominio como bloqueo para la preparación: los artefactos de despliegue se dejarán parametrizados. Bucket R2 hyt-documentos y claves continúan pendientes del usuario. Objetivo 7/8/9 activo.
## Dashboard visual del inicio — 03/11/2026

HomeController ahora consulta IReportingRepository solo con REPORTES.CONSULTAR y valida fechas Perú. Perfiles sin permiso conservan bienvenida y módulos autorizados, sin consultar ni recibir indicadores. DashboardModel/_Dashboard muestra cantidades por fuente autorizada, OC sin confirmar, listas para marcar, importes separados PEN/USD y enlaces a listados con fechas conservadas. No mezcla documentos como si fueran servicios distintos.

Gráficos SVG nativos: curvas de registros diarios con interpolación monótona entre puntos, columnas por estado y dona de tres categorías disjuntas de OC activas. Selección de series, foco y títulos de puntos, selector de intervalo utilizable en móvil y detalle tabular accesible. Rangos largos se agrupan en hasta noventa intervalos para acotar recursos. Días sin registros se incluyen con cero; sin series o datos se muestra estado vacío. Colores HYT rojo/negro/gris y CSS adaptable; sin dependencia de librerías externas ni animaciones obligatorias.

Pruebas HTTP nuevas aprobadas: inicio conectado, importes por moneda, tres gráficos y detalle accesible, rechazo de fechas invertidas y ausencia de consulta del repositorio para perfiles sin reportes. Chrome verificó dos curvas, ocultar/restaurar series, estado sin selección, detalle por fecha coincidente, barras/dona, rangos largos acotados y escritorio/móvil sin desbordamiento. Capturas dashboard-escritorio.png y dashboard-movil.png en entregables/preview revisadas visualmente; cifras de esas capturas son exclusivamente de repositorio de pruebas. Se mejoraron etiquetas móviles y geometría de dona tras revisión. Host temporal cerrado.

Pendientes del objetivo: casos adicionales de datos vacíos/parciales y permisos de fuentes, validación visual/impresión de reportes, cierre de pruebas/paquetes etapa7/8 y preparación parametrizada de publicación etapa9. El usuario indicó expresamente no desplegar aún: faltan claves y pruebas manuales. Objetivo activo.
## Preparación de etapa 9, sin publicación — 03/11/2026

Añadidos Dockerfile/.dockerignore, deploy/compose.hyt.yml, plantillas de parámetros privados/no secretos y Nginx. Servicio independiente en loopback, redes externas verificables para SQL/salida, volumen persistente de claves y ejecución no root. Dominio/IP/redes son parámetros pendientes. Guía GUIA_PUBLICACION_Y_QA.md incluye preparación, aceptación manual, respaldo SQL/objetos/claves, restauración aislada y reversión de versión.

Program configura encabezados reenviados solo desde IP explícita de proxy, con límite de una etapa y simetría IP/protocolo; exige IP válida en producción. Middleware anterior a redirección HTTPS. GET /health/live público devuelve solo ok como sonda de vida, sin comprobar ni exponer SQL/Cloudflare. Producción conserva DataProtection persistente requerido y Turnstile obligatorio.

Verificado: dotnet publish Release aprobado, sin archivos secrets.json/.env/certificados privados en salida; Docker Compose config aprobado sin resolver secretos; pruebas HTTP/C# de sonda y confianza limitada de proxy aprobadas. Motor Docker local no disponible: no se construyó ni ejecutó imagen y no se inició Docker Desktop. Activación, claves, bucket y pruebas manuales siguen diferidas expresamente por el usuario. Objetivo continúa activo: cerrar pruebas y paquetes 7/8 y revisar preparación completa de 9.

## Estado consolidado — 05/11/2026

Etapa 7 implementada: marcado externo Facturado con datos 01/03, requisitos, corrección/retiro auditados, historial y permisos de visibilidad. Etapa 8 implementada: dashboard conectado con filtros comerciales, series de primeros eventos, día/semana/mes, columnas navegables, dona, pendientes históricos paginados y cuatro reportes HTML. Etapa 9 queda preparada conforme a la indicación del usuario; publicación y aceptación real se mantienen pendientes de claves, bucket y pruebas manuales.

Cierre técnico local: 136 comprobaciones SQL, 186 HTTP/C#, 148 fuentes validadas y cinco regresiones negativas. Los 82 SP instalados coinciden con las fuentes. Chrome y cuatro impresiones A4 verificadas; bandejas móviles legibles mediante desplazamiento horizontal y filtros uniformes. Paquetes incrementales y SHA256 actualizados. Detalle de cobertura, comandos y límites en VERIFICACION_ETAPAS_7_8_Y_9.md. No equivale a publicación ni aceptación productiva.
