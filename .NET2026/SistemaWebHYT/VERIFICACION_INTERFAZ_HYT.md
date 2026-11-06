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
