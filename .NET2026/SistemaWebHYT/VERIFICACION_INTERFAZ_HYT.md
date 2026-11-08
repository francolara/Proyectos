# Verificación de la interfaz unificada HYT

Fecha: 05/11/2026. Autor: FRANCO LARA.

## Entrega

Implementados los ocho puntos de PLAN_MEJORAS_INTERFAZ_HYT.md en una misma entrega local. Se mantienen MVC, ADO.NET y los procedimientos almacenados existentes. La compilación normal de Visual Studio y la salida Release preparada están actualizadas; no se realizó despliegue remoto.

| Requisito | Implementación | Evidencia |
|---|---|---|
| Acciones como botones | _Grid y acciones de tablas auxiliares, adjuntos, revisiones, consultas e impresiones conservadas usan botones visuales con espaciado y colores HYT. Conservan sus rutas, permisos y confirmaciones. | Pruebas HTTP y navegador de acciones de los listados; auditoría de las vistas operativas sin enlaces de acción simples. |
| Retirar historial de edición | Historial general del maestro y presupuesto visible solo en consulta. Configuración, componente de formulario compartido y marcado de facturado dejan de incluir el historial general en edición. Se conservan las revisiones, referencias documentales y consultas históricas funcionales. | Matriz de edición de cliente, material, presupuesto, empresa, guía y facturado en ambos temas; 165 comprobaciones SQL, incluidos eventos de auditoría, aprobadas. |
| Fechas y responsables | _TraceSummary compartido, con usuario/fecha de registro y usuario/fecha de actualización; cuatro columnas en escritorio, dos en tablet y una en móvil estrecho. Formato Perú y «Sin actualizaciones» cuando corresponda. | Contraste y geometría comprobados en los formularios; capturas de material en escritorio, tablet y móvil. |
| Popups destacados | Barra roja, título blanco, cierre visible, cuerpo desplazable y pie separado. También calendario y tarifas embebidas; el título del editor de tarifas se muestra una sola vez en la cabecera exterior. | Ayudas y ventanas documentales verificadas en ambos temas; apertura/cierre/Escape/foco, sincronización de iframe y calendarios dentro de ventana. |
| Títulos de vistas | Encabezado diferenciado sobre superficie del tema con borde rojo y título destacado. Código interno conservado bajo «Volver al listado», con contraste en ambos temas. | Verificación de encabezados y código interno en la matriz de navegador; capturas claras/oscuras. |
| Listados al ancho disponible | Vistas principales declaran Wide; layout elimina la limitación central para ellas y conserva el ancho de formularios. La tabla desplaza horizontalmente dentro de su contenedor en móvil. | Diez listados comprobados a 1920 px; recorridos a 768 y 390 px sin desbordamiento de página. |
| Descripción de material | LongText en metadatos y textarea con rows=3, fila completa y máximo 1500. Redimensionamiento vertical y valor multilineal conservado al guardar. | Medición real del control y verificación del texto enviado al repositorio, con normalización de los finales de línea nativos del formulario. |
| Modo oscuro | Selector en menú visible en móvil, preferencia hyt.theme guardada en navegador y preferencia del sistema como valor inicial. Colores semánticos para campos, tablas, trazabilidad, carga, gráficos y popups. Sincronización de ventanas del mismo origen. | Prueba de selección/persistencia, tema inicial del sistema, contraste de texto de al menos 4.5:1 en los elementos comprobados y sincronización de tarifas. Impresión con fondo blanco. |

## Pruebas realizadas

- Compilación de solución Debug y publicación Release sin errores ni advertencias.
- Suite HTTP/C# existente aprobada con las referencias visuales actualizadas. El enlace a la ficha documental conserva acceso a estados aunque su texto ya no mencione el historial.
- check_interface_browser.mjs: 255 comprobaciones de presentación, temas, tamaños, controles, contraste, popups, guardado multilineal e impresión. Resultado detallado: entregables/preview/verificacion-interfaz.json.
- check_maintenance_browser.mjs: ayudas, zonas entre páginas, predeterminados, moneda opcional, tasas y calendario aprobados.
- check_budget_browser.mjs: búsqueda de cliente/tienda, tarifa, cálculo, USD pendiente, guardado y consulta posterior aprobados.
- check_followup_browser.mjs: ventanas OC/guía/confirmación, carga/recuperación PDF privada, aceptación, regularización y marcado facturado aprobados.
- check_report_browser.mjs y check_report_pdfs.py: siete casos de ambos formatos de presupuesto, documentos de una y varias páginas, vacío, USD y descripción larga. A4, texto completo y contenido dentro del papel.
- check_captcha_loading_browser.mjs: respaldo del captcha, bloqueo de Turnstile, renovación, errores, doble envío, solicitudes simultáneas, respuesta demorada, cancelación, XHR y retorno de navegación aprobados.
- SQL aislado: 165 comprobaciones aprobadas; la base temporal se eliminó. No se insertaron datos de prueba en SistemaWebHYT de trabajo.
- Comparación local: los 83 procedimientos coinciden con las fuentes. Validador offline: 153 scripts, 48 tablas, 108 relaciones. Esta entrega no altera estructura ni procedimientos SQL.

Las esperas de los scripts de navegador de presupuesto/seguimiento/carga se ajustaron para esperar la carga del documento después de una navegación, evitando consultar un body transitoriamente inexistente. Las verificaciones del resultado funcional siguen activas.

## Evidencias visuales

En entregables/preview:

- interfaz-listado-light.png e interfaz-listado-dark.png.
- interfaz-material-light.png e interfaz-material-dark.png.
- interfaz-material-light-768.png, interfaz-material-dark-768.png y variantes de 390 px.
- interfaz-ayuda-light.png e interfaz-ayuda-dark.png.
- interfaz-documento-light.png e interfaz-documento-dark.png.
- interfaz-calendario-light.png e interfaz-calendario-dark.png.
- interfaz-dashboard-light.png e interfaz-dashboard-dark.png.
- nuevo-una.pdf, anterior-una.pdf y los demás casos del inventario report-validation.json.

Las capturas y PDF utilizan exclusivamente datos ficticios del repositorio de pruebas. La aplicación de trabajo conserva sus datos y credenciales.

## Archivos y comportamiento compartido

hyt-interface.css define los tokens semánticos de ambos temas y los componentes unificados. hyt-theme-switch.js se carga antes de los estilos para seleccionar el tema inicial y sincroniza el iframe mediante origen y emisor validados. _TraceSummary.cshtml/TraceSummary.cs concentran las fechas y responsables; _Trace.cshtml adapta el contrato existente a la nueva presentación. Los estilos anteriores reutilizan los tokens en lugar de imponer fondos blancos o texto negro en modo oscuro.

El calendario tiene una cabecera propia, altura limitada y posición contenida en la ventana. Los elementos con hidden siguen ocultos aunque compartan estilos de botón. Los documentos HTML autónomos de impresión mantienen sus estilos propios y los colores de impresión claros.

Para revisar la entrega local, iniciar GestionServicios.Web con F5. El selector claro/oscuro se encuentra en la barra del menú.

## Corrección: historial oculto también en consulta — 05/11/2026

Retirados los bloques de historial general de las consultas de Master, Budget y Followup. La nueva instrucción del usuario amplía el alcance anterior de solo edición. Suite HTTP aprobada con pruebas explícitas de ausencia del historial y presencia de Fechas y responsables para clientes, materiales, versión de tarifario embebida, presupuesto, guía y OC. Resultado: entregables/preview/check-historial-oculto-http.log. Auditoría interna conservada sin cambios SQL.

Actualización 05/11/2026: la ventana de tarifas del material se sustituye por **Agregar tarifario**, con IGV, PEN, USD opcional y estado; se retiran los pasos de versiones/publicación del recorrido del material. Alta y edición aprobadas en navegador; regresión de interfaz: 255 comprobaciones aprobadas. Ver CODIGOS_Y_TARIFAS_DESDE_MATERIALES.md.

Actualización 05/11/2026 — direcciones y presupuesto: combos dependientes de ubigeo en clientes/tiendas; carga automática de direcciones; ayudas de Zona informativa/Especialidad; cuatro trimestres; retiro visual de Encargado/Firmante/Coordinador/Cargo del firmante; Consideraciones al final con predeterminado editable desde Configuración HYT. Recorrido Chrome aprobado y regresión general de 255 comprobaciones aprobada. Evidencias ficticias: cliente-ubigeo-escritorio.png, presupuesto-direcciones-ayudas.png y presupuesto-direcciones-movil.png, en entregables/preview.

Actualización 05/11/2026 — Detalle del presupuesto al ancho completo, Agregar ítem rojo y ayudas por fila. Chrome aprobó unidad automática, tarifa limitada al material, precio neto editable, necesidad/servicio, subtotal y cambio de moneda con importe ausente. Pruebas SQL: 190 comprobaciones aprobadas. Se sustituyen los controles de origen/IGV de cada fila por precio neto y se mantienen las referencias y auditoría internas.

06/11/2026: lupas de 34 px, necesidad de 46 px, tacho accesible, Ver/Editar de 84 px y tarjetas con borde superior rojo. Probada la apertura de tarifa sin cliente en Chrome y el número visible en la ficha de presupuesto; comprobaciones SQL de numeración mensual, concurrencia y repetición de migración aprobadas (195 totales).

06/11/2026: detalle con columnas proporcionales (material 40 %), sin ancho vacío dentro de su celda. Espacio restante distribuido entre SAP, servicio, necesidad y valores; botones de lupa/tacho y campos alineados a 44 px. Chrome verificó distribución y alturas en escritorio y ausencia de desborde exterior en 1920/1440/768/390 px. Evidencia: entregables/preview/detalle-espacio-redistribuido.png. Sin cambios de SQL ni lógica de guardado.

06/11/2026: acceso sin selector manual cuando Turnstile funciona. Respaldo automático ante fallo o ausencia de carga; la interacción normal cancela la espera de arranque y la recuperación restablece Turnstile. Cerrar sesión y Volver al inicio usan botones con iconos. Compilación y pruebas HTTP aprobadas; Chrome verificó captcha, espera interactiva, recuperación, loading y formulario de salida con POST protegido. Evidencia: entregables/preview/cerrar-sesion-botones.png.

06/11/2026: estados comerciales Registrado/Enviado/Observado/Aceptada/Rechazada/Cancelado; Vencido retirado como destino. Cambio de estado desde listado mediante popup HYT, motivo solo al observar/rechazar/cancelar y fecha/usuario automáticos. Observado editable en la misma revisión; lista sin cuatro columnas de auditoría y consulta conserva Fechas y responsables. Pruebas HTTP aprobadas, 208 comprobaciones SQL aisladas y recorrido Chrome en escritorio/móvil aprobados. Evidencias: presupuesto-cambiar-estado-listado.png y presupuesto-estado-movil.png. Ambos SP actualizados en SQL local y 84 módulos coinciden con fuentes.

06/11/2026: todos los botones de acción de tablas (incluidos Ver, Editar y acciones documentales) comparten 168 × 48 px y una rejilla de dos columnas. Verificado en Chrome para Presupuestos, Clientes, OC y Tarifarios, sin texto cortado ni desborde exterior móvil; conserva temas y colores. Evidencia: entregables/preview/acciones-presupuesto-uniformes.png.

06/11/2026: se retiran Revisiones del presupuesto, Crear nueva revisión y Registrar cambio de estado de la consulta/edición del presupuesto. El popup de cambio de estado queda exclusivamente en el listado. Se conservan revisiones, auditoría y datos de responsables internos.

06/11/2026: comparación visual con hytpresupuesto.pdf y reconstrucción del formato nuevo sin campos internos, tabla roja y subtotal amarillo. Configuración HYT permite firma JPG/PNG privada R2 con vista previa; imagen embebida e inmutable en nuevas impresiones. Pruebas HTTP y 213 comprobaciones SQL aprobadas; siete escenarios PDF en Chrome (ambos formatos, una/varias páginas, vacío, USD y descripción extensa) sin pérdida de texto ni desborde de papel. Evidencias nuevas: entregables/preview/firma-presupuesto/. Mapeo y campos pendientes: MAPEO_FORMATO_NUEVO_PRESUPUESTO.md. SQL local: 85 SP coinciden con fuentes. La firma de las pruebas es solo un recurso ficticio; la empresa debe subir su imagen real.

06/11/2026: Posición imprime el código interno del material; SQL toma Centro SAP FICO desde la tienda. Subtítulo de la cotización editable en Configuración HYT, precargado con el texto completo indicado y centrado en ambos formatos. HTTP, Chrome y siete PDF aprobados; 218 comprobaciones SQL aisladas aprobadas. Evidencias: entregables/preview/subtitulo-cotizacion/. Fuentes y 85 SP locales coinciden.

06/11/2026: cabecera del formato nuevo con Fecha 40 % más estrecha y número completo en una línea; cuadros informativos de 39 mm (30 % más anchos), centrados y con altura uniforme de 5 mm. Descripción del trabajo retirada de ambos formatos, conservando las descripciones de materiales. Versión HYT-2026.5.

Verificación de cuadros 06/11/2026: Chrome confirmó ancho de valores de 39 mm, alturas iguales, contenido centrado, Fecha al 9,3 % del ancho y espacio suficiente para número de doce dígitos en una sola línea. Siete PDF A4 aprobados y descripción interna ausente en ambos formatos. Evidencias: entregables/preview/cuadros-presupuesto/.

06/11/2026: Configuración HYT agrupa Subtítulo del presupuesto, Firmante predeterminado, Cargo del firmante y Consideraciones predeterminadas en tarjeta Configuración de presupuesto, con borde rojo y distribución adaptable. Se conserva el formulario y guardado conjunto.

06/11/2026: sección Subir firma para presupuesto con cabecera compacta de 16 px, fondo suave y borde superior rojo, acorde a Configuración de presupuesto. Archivo y botón en una fila en escritorio; disposición vertical en móvil.

06/11/2026: formato nuevo con Tienda al 9,75 % (reducción de 25 % sobre su ancho anterior de 13 %) y Centro SAP FICO al 10,25 % (antes 7 %). El resto de columnas conserva su ancho y el total es 100 %. Plantilla HYT-2026.6.

06/11/2026: corregido Centro SAP FICO vacío en impresiones nuevas de presupuestos anteriores; Obtener usa tienda como respaldo solo cuando el ítem no conserva valor. Presupuesto local 202611000001 tenía el detalle NULL y su tienda sí tenía centro configurado. No se reescriben partidas ni HTML guardados.

Verificación FICO 06/11/2026: 221 comprobaciones SQL aisladas aprobadas. Consulta real del presupuesto 202611000001 devuelve 9800677867HG en todas las partidas, usando la tienda como respaldo del valor antes vacío. Los 85 procedimientos instalados coinciden con fuentes.

06/11/2026: tipografía Arial en formato nuevo conservando anchos, cuadros y distribución. Subtítulo centrado en Arial regular de 6 pt con línea roja superior, conforme a referencia. Plantilla HYT-2026.7.

Verificación tipográfica 06/11/2026: Chrome confirmó Arial en cuerpo y detalle, subtítulo regular centrado y línea roja superior. Siete PDF A4 aprobados, incluidos varios ítems y descripción larga sin truncamiento. Evidencias: entregables/preview/tipografia-presupuesto/.

06/11/2026: formato ANTERIOR reconstruido contra el Excel original, conservando siete columnas y costo final sin IGV rosado, con Arial, subtítulo rojo centrado y firma común. Diez escenarios Chrome/PDF A4 aprobados (ambos formatos, varias filas, vacíos, USD y descripciones largas), sin pérdidas ni texto fuera del papel. Neto/IGV/total en revisión antigua CON_IGV verificados. Evidencias: entregables/preview/formato-anterior-adaptado/.

Subtarjeta de impresión dentro del encabezado Presupuesto verificada en consulta y edición guardadas, escritorio y móvil; ausente antes de guardar o sin permiso de reportes. Dos clics reales generan cada formato en pestaña nueva. Enlace de impresiones conservadas retirado; persistencia y permisos conservados. Pruebas HTTP y compilación aprobadas.

06/11/2026: formato anterior con una sola fila de resumen, COSTO FINAL SIN IGV; retiradas filas IGV y TOTAL. Tienda muestra únicamente su descripción. Sin cambios del formato nuevo ni del cálculo guardado.

Verificado en Chrome: resumen anterior de una sola fila y Tienda Centro sin prefijo T001; diez PDF A4 aprobados, incluido USD con costo neto de $ 100.00. Compilación y pruebas HTTP aprobadas. Evidencias: entregables/preview/anterior-sin-igv-total/.

06/11/2026: formato anterior muestra Coordinador/Encargado: bajo Tienda con los nombres de contactos activos asociados a esa tienda. Si hay varios, se separan con /; sin contactos, no se inventa nombre ni se toma uno general del cliente.

Verificado 06/11/2026: 224 comprobaciones SQL aprobadas (sin contacto, activo, inactivo, general, otra tienda y varios nombres), pruebas HTTP y diez PDF aprobados. Consulta real del presupuesto 202611000001 recupera el contacto de su tienda. Evidencias: entregables/preview/anterior-contacto-tienda/.

06/11/2026: captura documental con número de presupuesto a la derecha en tarjeta de borde rojo. OC sustituye selector Revisión del presupuesto por Estado del presupuesto informativo con nombres comerciales. La revisión asociada se resuelve en el servidor y conserva control de versiones.

06/11/2026: acciones de grillas y ayudas alineadas a la derecha; columna ajustada al contenido y grupos de una sola acción sin reservar una segunda columna vacía. Verificación en navegador de clientes, tiendas, especialidades, presupuestos, OC y guías en escritorio/móvil; ayuda de selección de cliente. Se conserva el tamaño uniforme y el detalle editable del presupuesto.

06/11/2026: OC compacta sin columnas Fecha comprobante/Tipo/Serie/Número. Doble clic o Enter en Facturado Sí abre consulta. Adjuntos OC/confirmación/Guía en pop-ups de solo consulta con fecha de carga y descarga privada. Nueva columna Guía y Adjuntos de guía. Verificación HTTP de permisos/consulta, 230 comprobaciones SQL y navegador escritorio/móvil con tools/check_oc_detalles_browser.mjs (host --browser-host --browser-oc-details). Nuevas vistas Razor incluidas automáticamente por el SDK web.

06/11/2026: OC reemplaza Ver por Cambiar estado en pop-up con estado actual y Fechas y responsables. Guardado AJAX conserva los filtros locales, versiones originales, permisos y validaciones SQL; fallos se muestran dentro del formulario. Presupuesto y Guía se presentan en tarjetas con borde superior rojo y enlaces normales a sus registros según permisos de consulta. tools/check_oc_estado_browser.mjs verifica navegación de ambas tarjetas, captura/guardado y retorno filtrado en escritorio/móvil. La vista nueva se incluye automáticamente mediante el SDK web.

06/11/2026: botones Guía, OC del cliente y Ejecución de presupuestos con borde rojo de 3 px y marca de registro si existe su ID. Abren consulta documental con detalle y acceso a visualizar/descargar archivos privados disponibles; altas sin registro conservan pop-up y estilo normal. Consulta SQL añade EjecucionId al final para ADO.NET. Verificados permisos HTTP, 231 comprobaciones SQL y navegador escritorio/móvil mediante tools/check_presupuesto_documentos_browser.mjs (host --browser-host --browser-budget-documents).

06/11/2026: OC retira Fecha confirmación y Adjuntos de confirmación del listado. Clic en el número de confirmación abre consulta de fecha, número, representante/observación, estado, responsables y archivos (fecha de carga, visualización/descarga), incluso si no tiene adjuntos. Ver confirmación usa la misma consulta. Compilación, suite HTTP y tools/check_oc_detalles_browser.mjs correctos en escritorio/móvil. Sin cambios SQL.

06/11/2026: OC registrada desde listado de presupuestos reutiliza la acción Edit y el mismo formulario pop-up del listado OC para usuarios con edición. Conserva borde rojo, adjuntos, responsables y motivo; guardar retorna a presupuestos con sus filtros. Consulta sin edición conserva ficha de lectura. Verificación HTTP y navegador con guardado/retorno correctos. Sin cambios SQL.

06/11/2026: corrige cierre involuntario al cancelar selección de archivo. El manejador cancel de hyt-ui.js procesa exclusivamente eventos cuyo target sea el propio diálogo; ignora los cancel propagados por input file. Se mantienen cierre deliberado y controles de formulario ocupado. tools/check_cancelar_adjuntos_browser.mjs verifica OC, Guía y confirmación sin selección/con archivo previo (datos y archivo conservados), cierre explícito del diálogo y firma HYT en configuración. Sin cambios de formularios, SQL ni archivos subidos.

06/11/2026: nuevas cargas privadas por carpeta de documento (OCCliente, Guia, ConfirmacionCliente), número seguro y GUID de operación; firma en ConfigHYT/Firma. Presupuestos/{número} disponible en generador para futura carga. C# y función SQL comparten saneamiento. Compatibilidad de referencias históricas y descarga/integridad preservadas. Suite HTTP/R2 simulado y 234 comprobaciones SQL correctas; script ACTUALIZAR_CARPETAS_R2.sql aplicado a base local. No se movieron objetos ni se subieron archivos de prueba al bucket real.

06/11/2026: se retira Facturado como columna del listado OC. El botón Confirmación del cliente usa borde rojo/marca si tiene ConfirmacionId vigente, y Facturado si existe referencia de comprobante vigente. Reutiliza estilo de documentos registrados de Presupuestos; conserva captura y consulta de los pop-ups. Compilación, suite HTTP y navegador escritorio/móvil correctos, incluyendo caso sin registros sin borde. Sin cambios SQL.

06/11/2026: listado OC retira Adjuntos de guía y concentra acceso en Guía: - sin GuiaId y Sí con registro. Clic abre consulta del registro (número, fechas, aceptación, representante, observación, estado y responsables) junto a archivos privados para visualizar/descargar. Acceso disponible con GUIAS.CONSULTAR, incluso sin adjuntos. Suite HTTP y navegador escritorio/móvil correctos. Sin cambios SQL.

06/11/2026: Guía sin estados; Ver reutiliza _Form bloqueado y Editar el mismo formulario habilitado. Consulta conserva PDF/descarga, responsables y enlaces del expediente. Sin Guardar, cargar, retirar ni motivo en consulta. Estado/filtro/transiciones retirados; web/SQL rechazan cambio de estado de Guía. Validaciones, trigger de Facturado, cierre OC y dashboard/reportes pasan a presencia de Guía, conservando ejecución terminada/confirmación. 230 comprobaciones SQL y suite HTTP correctas; navegador tools/check_guia_consulta_edicion_browser.mjs y check_oc_detalles_browser.mjs en escritorio/móvil. Script ACTUALIZAR_GUIA_SIN_ESTADOS.sql preparado.

## KPI y pendientes — 06/11/2026

- Compilación solución: cero errores/advertencias. Suite HTTP completa aprobada, incluidos permisos de facturación, seis/cuatro KPI, parámetros inválidos y filtros conservados al paginar.
- SQL real: 234 comprobaciones en base temporal eliminada; cohorte de 25 presupuestos y 21 OC verifica conteos completos, página 2 y combinación tipo/pendientes/estado.
- Navegador MVC aislado: 1440 y 390 px sin desbordamiento exterior, filtros GET y consulta/descarga de adjuntos Guía. Evidencias entregables/preview/kpi-*.png; validación reproducible tools/check_kpi_listados_browser.mjs.

## Confirmación completa OC — 06/11/2026

241 comprobaciones SQL reales aprobadas en base temporal eliminada: cambio automático sin ejecución, reintento sin duplicar auditoría, transición manual con confirmación, cierre con requisitos conservados, edición sin degradar OC cerrada y reversión integral ante fallo de auditoría. Evidencia entregables/preview/oc-confirmacion-automatica-sql.log.

## Facturado exige confirmación — 06/11/2026

SQL mantiene el bloqueo de altas/correcciones sin confirmación vigente de esa OC. Se verifica en la suite SQL (241 comprobaciones). El formulario informa el requisito, deshabilita la captura y omite Guardar referencia si no hay confirmación ni marcado previo. Referencias existentes conservan retiro auditado; SQL vuelve a validar cualquier guardado. HTTP completo aprobado con nuevo caso de formulario bloqueado. Compilación aislada cero errores/advertencias; la compilación habitual encontró archivos bloqueados por la aplicación de pruebas manuales en ejecución. Publicación local actualizada.

## Facturado con OC completada — 06/11/2026

243 comprobaciones SQL reales aprobadas en base temporal eliminada. Caso nuevo: OC completada automáticamente por confirmación, sin ejecución, falla sin guía y guarda Facturado después de registrar guía. Se mantienen pruebas de rechazo sin confirmación, idempotencia, corrección/retiro y auditoría. Evidencia entregables/preview/facturado-oc-completada-sql.log.

## Guía/Ejecución con ficha compartida en popup — 06/11/2026

Listado Guía: Ver abre _Form en consulta bloqueada con Editar registro según permiso; ese botón reemplaza el contenido del mismo diálogo por edición con Guardar registro. Editar abre directamente esa edición. Desde presupuesto, Guía registrada abre la misma consulta; Ejecución registrada abre edición (consulta si solo dispone del permiso de lectura), y sin registro usa alta existente en popup. Todos conservan returnUrl al listado de origen y la tarjeta de presupuesto. Details de Ejecución reutiliza también la ficha bloqueada: no muestra el panel antiguo de transición de estado. No se modifican estados históricos ni procedimientos SQL.

Compilación y suite HTTP completas aprobadas. Navegador MVC real aislado confirma Guía consulta→edición en mismo popup, Guía desde presupuesto, Ejecución editable sin estados, guardado y retorno a presupuesto. 1440/390 px sin desbordamiento exterior. Script tools/check_guia_ejecucion_popup_browser.mjs, evidencias entregables/preview/ejecucion-popup-*.png.

## Reportes integrales y dashboard vigente — 06/11/2026

- Compilación completa: cero errores y advertencias. 291 comprobaciones HTTP y 253 SQL reales en base temporal eliminada. SQL verifica todas las referencias vigentes, fecha de registro de relaciones, exportación mayor a 20, estados propios de OC, combinación de pendientes/tipo, permisos y conciliación de KPI con reportes. HTTP verifica catálogo de dos fuentes, filtros, paginación, exportación completa, validación y permisos.
- Navegador MVC real aislado: dos opciones de reporte, estado dependiente de fuente, ausencia de filtros Cliente/Tienda/Seguimiento, referencias, logo, impresión completa de 41 filas. Escritorio 1600px y móvil 390px sin desbordamiento exterior; modo oscuro verificado.
- Gráficos: curvas, dos columnas de estados independientes y dona dibujados con los estados actuales. Prueba reproducible tools/check_reportes_integrales_browser.mjs.
- PDF de QA: siete páginas A4 horizontales, 41 referencias de comprobante (todas las filas), logo y filtros en primera página, encabezados repetidos. Render revisado con Poppler. Evidencias entregables/preview/reporte-integral-horizontal.pdf, reporte-integral-pdf-validacion.json y reporte-integral-poppler-primera.png.
- Documentación funcional: FLUJO_FUNCIONAL_ACTUAL_HYT.md; técnica: DOCUMENTACION_BD_Y_FUNCIONALIDADES.md. Patch incremental ACTUALIZAR_REPORTES_DASHBOARD_VIGENTE.sql; paquetes previos sincronizados para no restaurar definiciones antiguas. Sin cambios de tablas.

Validación adicional de entrega: instalación aislada mediante paquetes incrementales (--use-bundles) y las 253 comprobaciones SQL aprobadas; base temporal eliminada. Procedimientos aplicados en SistemaWebHYT local y publicación local preparada.

## Menú Reportes abre consulta — 06/11/2026

Se separan acciones/rutas Index (/Reportes) y Print (/Reportes/Imprimir), compartiendo validación y consulta privada. Corrige generación ambigua de URL de RedirectToAction/Tag Helpers que enviaba el menú y Consultar reportes del dashboard a impresión. Pruebas HTTP verifican redirección /modulo/reportes → /Reportes y enlace de consulta del dashboard; impresión completa mantiene ruta propia.

## Dashboard con nombres legibles — 06/11/2026

Suite HTTP aprobada: controles y detalle sin Primera confirmación/Primer marcado. SQL real: 254 comprobaciones, incluyendo conciliación de registros presupuesto/OC y conteos vigentes de confirmación/Facturado sin historia retirada. Navegador verifica etiquetas exactas Confirmación del cliente/Facturado y que al activar ambas curvas el detalle usa los mismos nombres. Evidencia entregables/preview/dashboard-confirmacion-facturado.png; logs dashboard-nombres-http.log y dashboard-nombres-sql.log. UTF-8/CRLF conservados.

## Ajustes de Auditoría, acceso y Reportes — 07/11/2026

- Auditoría conserva la acción al consultar y agrupa los eventos del presupuesto y sus revisiones en **Presupuestos**. Se verificó con 311 comprobaciones HTTP/C# y 269 comprobaciones SQL en una base aislada. El procedimiento se actualizó en SQL local.
- Chrome verificó el control de visibilidad de contraseña en inicio de sesión, cambio de contraseña y alta de usuario. El cambio de contraseña mantiene un campo por fila en escritorio y móvil.
- Se verificó en la vista compartida que el listado y la impresión HTML de Reportes usan el encabezado **Guía**.
