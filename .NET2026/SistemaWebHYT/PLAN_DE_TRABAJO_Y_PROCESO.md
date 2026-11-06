# Sistema Web HYT: proceso y plan de trabajo

Fecha: 03/11/2026. Estado: propuesta técnica inicial para revisión.

## 1. Base del análisis

Fuentes: `Gestion_de_Servicios_Propuesta_Funcional.pdf`, versión 0.5 del 23/10/2026, y documento existente `GestionServiciosSupermercados/DOCUMENTACION_FUNCIONAL.md`, versión 0.6 del 24/10/2026. El segundo aclara que «Guia (Acta de conformidad)» y acta de conformidad son el mismo documento. Las propuestas del levantamiento no se consideran automáticamente reglas aprobadas.

Decisión confirmada por Franco: una sola empresa operadora, H Y T Servicios Generales E.I.R.L.; múltiples clientes, tiendas y usuarios. No habrá selector de empresa, suscripciones ni aislamiento multiempresa. La ficha de HYT será configuración única para identificación fiscal e impresión.

La base creada actualmente contiene plantillas de web y biblioteca Razor; todavía no implementa acceso, operaciones ni persistencia.

## 2. Proceso de atención por invitación

1. Comercial abre un expediente y registra cliente, tienda, solicitud, responsable y fecha real de recepción.
2. Selecciona un tarifario vigente del cliente o genérico y elabora el presupuesto. Guarda descripción, unidad y precio de cada partida como copia histórica.
3. Registra el envío de una revisión. Si el cliente pide cambios, crea otra revisión; conserva la anterior.
4. Registra aprobación con fecha, representante del cliente y evidencia. Si se rechaza, registra el rechazo sin continuar automáticamente.
5. Operaciones registra la OC recibida del cliente y la vincula con la revisión aprobada. No se emite una OC en nombre del cliente.
6. Programa, inicia y termina el trabajo, conservando fechas previstas y reales.
7. Registra manualmente la Guia (Acta de conformidad): cabecera, partidas, cantidades realizadas y archivo. Registrar el acta no significa que el cliente la aceptó.
8. Si el acta tiene observaciones, registra subsanación y posterior aceptación con evidencia.
9. Con OC registrada, trabajo terminado y acta aceptada, cierra operativamente la OC.
10. Gestiona la confirmación interna del cliente. Cada solicitud y respuesta conserva su propio registro.
11. Con confirmación válida, el expediente queda autorizado para facturar.
12. Marca la Orden de Compra - Cliente como Facturado y registra fecha, tipo de documento, serie y número del comprobante emitido fuera del sistema. La emisión de facturas no está incluida en esta etapa.

## 3. Proceso de emergencia

1. Se registra el presupuesto con tipo **Emergencia**, cliente, tienda y datos de la atención. Puede comenzar como borrador, sin valoración comercial completa y sin Orden de Compra - Cliente.
2. Se conserva la ejecución real y su fecha, aunque el trabajo haya ocurrido antes del registro del presupuesto en el sistema.
3. Se registra la **Guia (Acta de conformidad)** vinculada obligatoriamente al presupuesto, sin exigir Orden de Compra - Cliente. Se conserva su detalle manual y se gestiona la aceptación.
4. Se completa y envía el presupuesto de emergencia con su tarifario, sin sobrescribir el detalle propio de la guía.
5. Al recibir la OC emitida por el cliente se registra como **Orden de Compra - Cliente**, vinculándola al presupuesto y a los antecedentes existentes. No se reinicia el trabajo ni se duplica la guía.
6. Con Orden de Compra - Cliente, trabajo terminado y guía aceptada se registra cierre operativo.
7. Se gestiona la confirmación interna del cliente y después se habilita la facturación.

La aclaración del usuario del 03/11/2026 prevalece sobre el recorrido inicial del levantamiento: en el sistema la guía se registra contra un presupuesto existente, que puede estar en borrador para emergencia. Esto no exige que la cotización se haya enviado o aprobado antes de ejecutar el trabajo; las fechas reales se conservan por separado.

La aprobación comercial separada del presupuesto de emergencia queda pendiente de validación. No se impondrá como requisito confirmado.

## 4. Controles del proceso

| Acción | Control del servidor |
|---|---|
| Enviar revisión | Datos completos, partidas e importes válidos; conservar versión enviada |
| Iniciar invitación | OC registrada y programación definida |
| Registrar atención de emergencia | Presupuesto de tipo Emergencia con datos mínimos; no exige Orden de Compra - Cliente |
| Aceptar acta | Identificar representante, fecha y evidencia según regla acordada |
| Cerrar OC | Trabajo terminado y acta aceptada vinculados a la OC |
| Habilitar factura | OC, trabajo terminado, acta aceptada y confirmación vigente del cliente |
| Corregir o anular | Permiso, motivo e historial; conservar registros originales |

Los estados de presupuesto, ejecución, OC, acta y facturación son independientes. El resumen del expediente se calcula a partir de ellos, sin permitir un botón de «avanzar» que omita requisitos. Una eventual corrección de estado no debe habilitar facturación sin validar nuevamente sus condiciones.

Cada acción registra usuario y hora del sistema, separados de la fecha del evento y del documento. El almacenamiento se propone en UTC y la presentación en America/Lima. No inventar fechas para antecedentes desconocidos.

## 5. Pantallas y navegación

Acceso privado con recuperación de contraseña y usuarios administrados. Después del acceso se abre el panel, sin selección de empresa.

Menú lateral y barra superior siguiendo la organización de FralseCont; identidad de HYT, usuario y cierre de sesión. Menú plegable en celular y formularios adaptados al ancho disponible.

| Grupo | Pantallas |
|---|---|
| Inicio | Panel de pendientes y emergencias por regularizar |
| Maestros | Clientes, tiendas, contactos, servicios, materiales y unidades |
| Comercial | Tarifarios, presupuestos y revisiones |
| Operaciones | Expedientes, OC recibidas, ejecución y Guia (Acta de conformidad) |
| Seguimiento documental | Espera de confirmación y marcado Facturado con referencia del comprobante externo |
| Consultas | Histórico y reportes |
| Administración | Usuarios, permisos, configuración única de HYT y auditoría |

### Nombres de documentos y relaciones confirmadas

Aclaración del usuario del 03/11/2026:

- **Presupuesto**: registro inicial documental con campo obligatorio **Tipo de atención**: Emergencia o Invitación. El tipo se define en el presupuesto y se muestra en listado, ficha, ayudas y reportes. Cualquier cambio deberá auditarse y validar que no contradiga documentos o transiciones existentes.
- **Orden de Compra - Cliente**: denominación visible del registro de la OC emitida por el cliente; se relaciona con el presupuesto. No representa una compra a proveedores de HYT.
- **Guia (Acta de conformidad)**: denominación visible acordada para el documento manual de conformidad. Se vincula obligatoriamente al presupuesto; la relación con Orden de Compra - Cliente es opcional mientras no exista en una emergencia. Guardar o aceptar la guía no crea automáticamente una OC.

Para Invitación: Presupuesto → aprobación → Orden de Compra - Cliente → ejecución → Guia (Acta de conformidad). Para Emergencia: Presupuesto de emergencia registrado, incluso en borrador → Guia (Acta de conformidad) sin OC → completar presupuesto → recepción y registro de Orden de Compra - Cliente. La ejecución real de emergencia puede ser anterior a la captura documental. En ambos casos se conserva la confirmación del cliente previa a facturar.

Los nombres visibles se aplicarán a menús, títulos, ayudas, listados, reportes e impresiones. Cada presupuesto admite como máximo una Orden de Compra - Cliente. Cada presupuesto admite una sola Guia (Acta de conformidad) y corresponde a una tienda específica, obligatoria en su cabecera. La prueba de emergencia debe permitir guardar una guía con presupuesto y sin OC, pero rechazar una guía sin presupuesto.

### Identificación interna y OC emitida por el cliente

Requisito confirmado el 03/11/2026: Orden de Compra - Cliente tendrá dos campos separados: **N° de registro interno**, asignado por HYT, y **N° OC del cliente**, ingresado desde el documento recibido. El número externo no reemplaza al interno. Se conservará como texto alfanumérico, incluidos ceros iniciales y separadores.

La ficha permitirá ingresar y consultar ambos. La grilla, búsquedas, ayudas y reportes los identificarán con etiquetas diferentes; el vínculo con presupuesto, guía y confirmación usará el identificador técnico del registro, no el número externo como clave. La fecha de emisión de la OC y la fecha de recepción también permanecerán separadas de la fecha automática de registro. La política de duplicados externos se validará por cliente, sin imponer unicidad global.

Los dos Excel de presupuesto son la referencia de campos para la definición de formularios e impresiones, con inventario en `ANALISIS_FORMATOS_PRESUPUESTO.md`. Conservar campos de ambos formatos; no eliminar ni unificar las referencias «PPTO N°» y «N° COTIZACION» sin aclarar su significado. La definición detallada puede avanzar con esta base; siguen pendientes las reglas ya identificadas, como IGV y cardinalidades documentales.

### Archivo de la Orden de Compra - Cliente

Requisito confirmado el 03/11/2026: la ficha de **Orden de Compra - Cliente** permitirá subir el archivo de la OC emitida por el cliente, además de registrar número interno y número externo. El archivo quedará vinculado a esa OC y accesible posteriormente mediante **Ver OC adjunta** y **Descargar**, desde su ficha y desde el indicador de adjunto del listado. Los PDF podrán abrirse en un visor o pestaña autorizada; otros formatos admitidos tendrán descarga cuando no sea posible previsualizarlos. Tipos permitidos y tamaño máximo se definirán antes de implementar la carga.

La carga será al almacenamiento privado de HYT; es independiente del enlace al portal externo configurado en el maestro de clientes. Subir el archivo no cambia por sí solo el estado de aceptación del cliente ni el marcado Facturado. Registrar usuario y fecha de carga; cualquier reemplazo o retiro conservará auditoría. Validar permisos al cargar, visualizar y descargar. La prueba incluirá subir, guardar, recargar y volver a consultar el archivo de la OC correcta, sin perder las dos numeraciones.

### Registro de Guia (Acta de conformidad) desde presupuestos

Requisito confirmado el 03/11/2026: el listado de presupuestos tendrá la acción **Registrar Guia (Acta de conformidad)**. Se usará un botón que abre una ventana emergente con el estilo común de FralseCont y colores HYT. En la grilla se mostrará un indicador de guía registrada y su estado; el indicador no guardará cambios por sí solo.

La ventana mostrará el presupuesto seleccionado y permitirá ingresar **fecha del acta**, **número de acta**, **observación** y **archivo escaneado PDF**. Fecha y número se proponen obligatorios; observación y PDF son opcionales, pudiendo registrar uno o ambos. El número se tratará como texto para conservar letras, guiones y ceros iniciales. Los datos obligatorios finales se validarán con el negocio.

Al guardar se vincula el registro al presupuesto y se actualizan sus indicadores sin perder la página y filtros del listado. En emergencias no se exige Orden de Compra - Cliente. Si existe una guía, se ofrece **Ver Guia (Acta de conformidad)** y su gestión autorizada, evitando crear duplicados por repetir el clic. Solo se permite una guía por presupuesto; si existe, se consulta o modifica con permisos y auditoría, sin crear otra.

Registrar una guía o adjuntar un PDF no registra automáticamente aceptación del cliente ni autorización para facturar. La aceptación conservará su acción y trazabilidad propias. La fecha del acta es distinta de la fecha/hora de registro asignada por el servidor.

El PDF se guardará como adjunto privado ligado a la guía, con validación de contenido y tamaño permitido. Consulta y descarga requerirán permiso; un reemplazo conservará su trazabilidad. Crear, editar, aceptar, anular y cambiar el adjunto registrará usuario, fecha y cambios; las transiciones incluyen estado anterior y nuevo. El listado mostrará número, fecha y estado de la guía, e indicador de PDF; la ficha mostrará observación, archivo e historial.

Esta acción es el acceso directo desde presupuestos a la guía. No impone volver a capturar todas las partidas en ese popup; el alcance del detalle manual de la guía se mantiene por validar antes de implementar el formulario completo. Pruebas: guardar sin OC en emergencia, cancelar sin cambios, consultar PDF autorizado, conservar datos tras error y evitar registros duplicados por reenvío.

### Confirmación interna del cliente desde Orden de Compra - Cliente

Requisito confirmado el 03/11/2026: cada fila del listado **Orden de Compra - Cliente** tendrá la acción **Registrar confirmación del cliente**, que abre el popup común con la OC y el presupuesto identificados. Permitirá registrar **fecha de confirmación**, **número de confirmación alfanumérico**, **observación** y adjunto opcional. El número de **Guia (Acta de conformidad)** también será alfanumérico. Ambos números se guardarán como texto, conservando letras, ceros iniciales y separadores, sin conversión numérica.

Al guardar, la grilla de OC mostrará confirmación registrada, fecha, número e indicador de adjunto; permitirá consultar los datos desde **Ver confirmación**. Se conservarán página y filtros. La fecha informada es la del evento y no sustituye la fecha/hora automática de registro. El guardado y las modificaciones tendrán auditoría, con estado anterior/nuevo cuando impliquen una transición. Registrar la confirmación no emitirá una factura ni omitirá validar los demás requisitos para habilitarla.

**Visualización posterior de adjuntos:** tanto la guía desde presupuestos como la confirmación desde OC ofrecerán **Ver adjunto** en su consulta/popup y desde el indicador del listado. Los PDF se visualizarán dentro de un visor o en una pestaña nueva con acceso autorizado; se permitirá descarga como alternativa cuando el navegador no pueda mostrarlos. Si hay varios adjuntos se mostrará su relación para elegir. No mostrar acciones de archivo cuando no exista adjunto. Mantener almacenamiento privado, permisos y trazabilidad de reemplazos; no exponer rutas internas ni enlaces públicos permanentes.

Pruebas: conservar números como `CONF-001A` y `GUIA-0007B`; guardar y volver a consultar fecha/observación; visualizar un PDF de guía y otro de confirmación después de recargar; impedir acceso a archivos sin permiso; cancelar sin cambios y evitar duplicados por reenvío.

### Marcado Facturado: alcance inicial sin emisión de facturas

Decisión confirmada el 03/11/2026: la facturación como emisión de comprobantes no entra en esta etapa. Se incluye únicamente seguimiento del documento emitido externamente mediante la acción **Marcar como facturado** en el listado de Orden de Compra - Cliente.

El check o botón abrirá un popup con **fecha del comprobante**, **tipo de documento**, **serie** y **número**. Se guardan los cuatro datos junto con el marcado, sin permitir un check aislado sin su referencia documental. El tipo se elegirá de un catálogo interno sin mantenimiento: Factura (01) o Boleta (03), únicos tipos incluidos por ahora. Se guardará directamente el código textual 01 o 03, sin identificador numérico y sin unicidad en los documentos relacionados. Serie y número se tratarán como texto, conservando letras y ceros iniciales. Cancelar deja el registro sin cambios.

La acción requiere que esté registrada la **aceptación/confirmación interna del cliente** de esa OC, capturada mediante la acción del apartado anterior; no basta la aprobación comercial del presupuesto. Se mantienen los controles documentales previos del flujo. Validarlos de nuevo en el servidor al guardar. Marcar Facturado no genera comprobante, numeración tributaria, envío a SUNAT ni asiento contable.

El listado y la ficha mostrarán Facturado, fecha, tipo, serie y número, además del usuario y fecha/hora de registro. Alta, corrección o retiro del marcado requieren permisos y auditoría; una reversión conserva referencia anterior y motivo. Evitar doble guardado por reenvío. La fecha del comprobante es independiente de la fecha de registro.

El reporte OC pendientes de factura utilizará este marcado para distinguir OC pendientes de confirmación, confirmadas pendientes de marcado y marcadas como facturadas. No se introduce en esta etapa cálculo de saldos por facturación parcial ni agrupación automática de OC; esas cardinalidades siguen pendientes de definición.

Pruebas: bloquear sin aceptación del cliente; confirmar con datos completos; cancelar sin marcar; conservar serie/número alfanuméricos; visualizar referencia tras recargar; auditar modificaciones y reversión.

### Listados, fichas y auditoría visible

Requisito confirmado el 03/11/2026: cada módulo de mantenimiento o documento tendrá un listado y una ficha con modos **Nuevo**, **Editar** y **Ver**. Ver será de solo lectura. Se conservarán filtros, orden y página al regresar al listado. Las acciones disponibles dependerán del estado y los permisos, con controles equivalentes en el servidor.

Se tomarán como referencia las grillas `workspace-table-wrap`, los listados y formularios de FralseCont y el componente `Views/Shared/_TrazabilidadRegistro.cshtml` de SistemaAdministrativoWeb. Se trasladarán los patrones a componentes compartidos de HYT, adaptando estilos y comportamiento mediante controladores y vistas parciales MVC; no se copiarán contexto multiempresa ni colores corporativos de FralseCont.

Los listados tendrán título, acción de alta cuando corresponda, filtros, búsqueda, ordenación, paginación, estado y acciones de consulta/edición. Se mostrarán fecha y usuario de registro, última actualización y último cambio de estado; las fechas relevantes del módulo, como aprobación, envío, aceptación, cierre o autorización, estarán disponibles en sus columnas y filtros. La presentación podrá agrupar fecha y usuario en una celda para mantener legibilidad. No se confunde fecha de creación del registro con fecha de emisión del documento.

La ficha incluirá un bloque común de **Trazabilidad del registro** con fecha/usuario de alta y última actualización, más los hitos propios del documento y un historial visible. La aprobación mostrará por separado quién registró la acción en HYT, quién aprobó por parte del cliente y cuándo ocurrió la aprobación. Campos no aplicables no se impondrán a todos los maestros; ausencia de una acción se mostrará como «Sin aprobación», «Sin actualizaciones» o equivalente, sin inventar fechas.

Cada transición de estado, incluidas aprobación, rechazo, observación, aceptación, cancelación, anulación, activación, desactivación, cierre y autorización, generará un evento con entidad/documento, estado anterior y nuevo, usuario, fecha y hora de registro, fecha del evento, motivo y referencia a evidencia cuando corresponda. Los eventos repetidos se conservarán íntegros. El resumen del listado refleja el último evento; la ficha permite revisar la secuencia completa.

El cambio y su evento de auditoría se guardarán de forma atómica: si falla la auditoría, no se confirma la transición. La fecha de registro la asigna el servidor y no se edita en formularios. Corregir una fecha del evento deja registro del valor anterior, nuevo, motivo y autor. Usuarios operativos no podrán editar ni borrar el historial. Las consultas estarán sujetas a permisos.

En móviles, filtros y acciones se reorganizarán; las grillas tendrán desplazamiento horizontal controlado y acceso al detalle sin recortar datos, con botones utilizables al tacto, etiquetas claras y navegación por teclado. Se probarán escritorio, tableta y celular. No ocultar la trazabilidad: si una columna no cabe en la vista inicial, debe seguir accesible y estar completa en la ficha.

### Ayudas emergentes y paginación

Requisito confirmado el 03/11/2026: las ayudas de materiales, zonas, unidades, clientes, tiendas y demás datos relacionados usarán el mismo patrón de ventana emergente de FralseCont. Referencia: `Views/PlanCuenta/Formulario.cshtml`, ayudas `workspace-modal`, cabecera, búsqueda, grilla desplazable y pie con acciones. Se adaptará a vistas parciales MVC y JavaScript común y colores HYT.

Cada ayuda tendrá botón de búsqueda junto al campo, título contextual, filtros aplicables, resultados, selección y acciones Confirmar/Cancelar. Al confirmar se devuelve identificador y descripción al formulario. Cancelar no modifica el valor previo. Para zonas, la ayuda permitirá selección múltiple, conservando las selecciones al buscar o cambiar de página y evitando duplicados. El usuario podrá revisar y quitar las zonas elegidas antes de confirmar.

Todos los listados y ayudas consultarán páginas de **20 registros desde SQL Server**; la última podrá contener menos. Búsqueda, filtros, ordenación y conteo se ejecutarán en el servidor. No cargar el catálogo completo para paginarlo en el navegador. Mostrar total de resultados, página actual y controles de navegación; cambiar filtros reinicia la página a 1. Validar selección y permisos nuevamente al guardar.

Las ventanas serán adaptables a celular, con cierre accesible, navegación de teclado, foco contenido y retorno de foco al campo que las abrió. Contemplarán carga, sin resultados y error con reintento conservando la búsqueda. La aceptación probará 0, 20 y 21 resultados, filtros y orden estables, cancelación y selección múltiple entre páginas.

### Búsqueda común y filtros de fechas

Requisito general confirmado el 03/11/2026: todos los mantenimientos tendrán buscador por **código o descripción/nombre**, sin exigir coincidencia exacta. Incluir los códigos aplicables a cada maestro; materiales permitirán buscar por código interno, posición, código SAP y descripción. Las ayudas emergentes conservarán el mismo criterio contextual.

Los registros documentales, incluidos **Presupuesto** y **Orden de Compra - Cliente**, tendrán además filtros **Desde/Hasta**, con **último mes** seleccionado al abrir el listado por primera vez. Criterio inicial explícito: mes móvil desde el mismo día del mes anterior hasta hoy en America/Lima; si no existe ese día, usar el último día de ese mes. No interpretar como mes calendario anterior. El usuario podrá ampliar o cambiar el rango; al volver de la ficha se recuperan sus filtros previos, sin restablecerlos automáticamente.

Los filtros se dispondrán en la barra común de FralseCont: buscador, fechas Desde/Hasta y acción Buscar, con adaptación móvil y colores HYT. Se mostrará la fecha utilizada: inicialmente **Fecha de registro** en listados documentales, diferenciada de emisión, aprobación y otras fechas del proceso. Los reportes podrán usar su fecha específica identificada. No imponer filtro temporal a los maestros, para evitar ocultar clientes o materiales antiguos.

Búsqueda y fechas se combinarán en SQL antes del conteo y la paginación de 20 registros. Cambiar o aplicar filtros vuelve a página 1; validar Desde menor o igual que Hasta. El día Hasta incluye sus registros hasta el final del día, sin perder operaciones por su hora. La aceptación probará código/descripción, búsquedas sin resultado, límites de fecha, retorno desde edición y filtros en móvil.

### Estilo común de fechas de FralseCont

Confirmado el 03/11/2026: todos los campos de fecha, incluidos filtros Desde/Hasta y popups de guía, confirmación y marcado Facturado, utilizarán el patrón de calendario de FralseCont. Referencias verificadas: `wwwroot/js/site.js` y clases `app-date-picker` de `wwwroot/css/site.css` en SistemaAdministrativoWeb.

Reproducir el campo de fecha con botón de calendario, panel emergente, navegación entre meses, días de semana y acciones **Hoy** y **Limpiar** cuando corresponda. Presentación de fechas **dd/MM/yyyy** y trazabilidad con hora **dd/MM/yyyy HH:mm**, siguiendo el componente `_TrazabilidadRegistro.cshtml`. Mantener etiquetas, validación, estado deshabilitado/solo lectura y funcionamiento dentro de popups y vistas móviles. Hoy usará la fecha de America/Lima; limpiar un campo obligatorio no evita su validación.

La implementación MVC conserva nombres y valores de los campos para model binding, con validación en controlador. El calendario adapta el JavaScript de FralseCont al tema HYT. Colores de selección y foco en rojo HYT, superficies blancas y texto negro. El formato visual no altera el tipo fecha ni el almacenamiento UTC de auditoría; fechas de documento sin hora no se desplazan de día por conversión de zona horaria. El control común de fechas está implementado como parcial MVC y JavaScript; los módulos lo utilizarán al desarrollarse.

### Identidad HYT y producto a medida

Este producto se construye a medida para H Y T Servicios Generales E.I.R.L., con sus nombres de proceso, formatos y reglas. La referencia funcional y de interacción es FralseCont; la identidad visual será de HYT. El logo existente en `GestionServiciosSupermercados/assets/logos/logo-cliente-hyt.png` contiene rojo, negro y blanco: se priorizarán estos colores, con fondos neutros para lectura y rojo como acento de marca. La paleta final deberá comprobar contraste y diferenciar acciones destructivas de acciones principales mediante texto e iconografía, sin depender solo del color. Conservar colores semánticos de estados cuando ayuden a entenderlos.

### Materiales, unidades de medida y zonas

Requisitos confirmados el 03/11/2026 tras recibir `NUEVO TARIFARIO  2026-2028.xlsx`:

- Incorporar mantenimientos de **Unidades de medida** y **Zonas**, ambos con listado, ficha, activación/desactivación y trazabilidad.
- El registro de material/servicio permitirá ingresar **posición**, **código interno** y **código de material SAP** como texto alfanumérico de libre registro. No exigir que sean números ni convertirlos a valores numéricos; conservar ceros iniciales. La generación automática de código interno no será obligatoria.
- Seleccionar una unidad de medida desde su mantenimiento y **una o varias zonas** desde el mantenimiento de zonas. No crear zonas mediante texto libre en el formulario de material. La selección múltiple se refiere a zonas; múltiples unidades por material no están solicitadas.
- Registrar descripción y estado del material. La posición y el código SAP no se asumirán únicos: el Excel contiene repeticiones con descripciones diferentes. Cada material tendrá un identificador técnico independiente. Validar la política de duplicados de código interno antes de cargar antecedentes.
- Mantener los precios, moneda y vigencia en la versión del tarifario, asociados al material y su alcance de zonas. Modificar una tarifa no cambiará presupuestos históricos.

En la etapa 3 se entregarán ambos mantenimientos y la selección múltiple de zonas en materiales. La prueba incluirá material con zonas Sur y Centro, recuperación de sus selecciones, uso de códigos alfanuméricos y conservación de referencias históricas al desactivar una unidad o zona.

El Excel contiene una columna de precio para el conjunto «Sur-Centro-Norte», no tres precios separados. Antes de importar se validará que ese precio aplica por igual a las tres zonas. Precios distintos por zona, si se requieren, se registrarán explícitamente sin inferirlos del archivo. Ver `ANALISIS_TARIFARIO.md`.

### Monedas y valores predeterminados del tarifario

Requisito confirmado el 03/11/2026: catálogo interno de monedas en base de datos, sin pantalla de mantenimiento. Registros iniciales: Soles, PEN, S/; Dólares, USD, $. La moneda inicial de presupuestos y tarifarios será PEN. En el presupuesto se podrá elegir PEN o USD desde ese catálogo. Las tablas relacionadas guardarán directamente el código textual PEN o USD, sin Id numérico; el campo de moneda podrá repetirse entre registros y no tendrá restricción de unicidad en esas tablas.

El tarifario tendrá un apartado de valores predeterminados: moneda (inicialmente PEN) y una o varias zonas. Cada registro/partida tendrá identificada su moneda y sus zonas, tomando esos valores al crearse y permitiendo particularizarlos por registro. Cambiar los valores predeterminados no reescribe partidas existentes ni documentos históricos. La selección múltiple de zonas se conserva.

En la ayuda de materiales del presupuesto, las zonas serán **informativas**. No constituyen filtro obligatorio, prioridad de resultados ni restricción de selección por ahora. Se mostrarán junto con código, descripción, unidad, moneda y precio. La paginación SQL de 20 registros se mantiene. Esta decisión sustituye cualquier interpretación anterior que use la zona para elegir o bloquear automáticamente una tarifa en el presupuesto.

Cada presupuesto conservará una sola moneda por revisión. Cambiar PEN por USD no se interpretará como conversión automática ni como simple sustitución del símbolo sobre importes existentes: la pantalla deberá advertir y requerir revisar/revalorizar las partidas antes de confirmar. El precio en dólares se registra manualmente y es opcional; no se convierte automáticamente mediante tipo de cambio. Ausencia de precio USD no equivale a cero ni autoriza usar el precio PEN como USD. No sumar precios de monedas distintas sin tratamiento explícito. Las impresiones nueva y anterior mostrarán la moneda real y no textos fijos de soles cuando el presupuesto esté en dólares.

Pruebas: presupuesto y tarifario nuevos en PEN; selección de USD; predeterminados aplicados a una partida nueva; cambio de predeterminados sin alterar registros anteriores; selección de material de otra zona sin bloqueo; cambio de moneda con partidas existentes sin conversión silenciosa.

### Portal de documentos por cliente

El maestro de clientes incluirá el campo **Enlace para subir documentos**, correspondiente al portal externo de cada cliente. El levantamiento lo identifica como «URL de carga de OC»; la aclaración del 03/11/2026 confirma su ubicación en la ficha del cliente y su finalidad de carga de documentos.

El usuario autorizado podrá registrar y editar la URL y abrirla con la acción **Abrir portal de documentos**, en una pestaña nueva. Se validará que sea una URL absoluta HTTP o HTTPS. Si no está registrada, se mostrará «Sin enlace registrado», sin habilitar la acción. La modificación conservará usuario y fecha en auditoría.

El enlace pertenece al cliente, no a cada tienda. Guardarlo o abrirlo no confirma que se hayan subido archivos; la carga se realiza en el portal externo. No se almacenan contraseñas del portal ni se presupone una integración automática.

La prueba de aceptación del maestro incluirá guardar enlaces distintos para dos clientes, recuperar el enlace correcto al consultar cada ficha y comprobar el estado sin enlace.

La ficha del expediente presenta resumen, presupuesto y revisiones, OC, ejecución, actas, confirmaciones, facturas, adjuntos e historial. Debe mostrar qué falta para continuar.

### Formato uniforme y reportes HTML

Decisión del usuario del 03/11/2026: mantener un mismo formato para todas las pantallas y generar los reportes en HTML. Se compartirán menú, barra superior, títulos, filtros, botones, formularios, tablas, validaciones y estados vacíos. Los reportes tendrán vista HTML de consulta y estilos de impresión; PDF mediante impresión del navegador como opción derivada. Excel queda como exportación adicional por definir, no como motor obligatorio del reporte.

Alcance confirmado por el usuario el 03/11/2026: implementar ambos formatos de impresión de presupuestos en HTML, **Formato nuevo** y **Formato anterior**, basados en los dos Excel proporcionados. El análisis está en `ANALISIS_FORMATOS_PRESUPUESTO.md`. Ambos son entregables obligatorios del plan; cuál se ofrecerá por defecto queda por definir.

La vista previa permitirá seleccionar el formato antes de imprimir o guardar como PDF desde el navegador. Las variantes conservarán sus campos y presentación documental, compartiendo estilos y componentes comunes. Elegir un formato no modificará las partidas ni recalculará el presupuesto con reglas diferentes. Si faltan datos necesarios para una variante, la vista previa los identificará para completarlos sin inventar valores.

Se conservarán la identificación y versión del formato utilizado al generar una impresión de una revisión del presupuesto, junto con usuario y fecha. Las pruebas cubrirán ambos formatos, una y varias páginas, descripciones largas, encabezados repetidos, totales y firmante. Se aplicará la selección CON IGV/SIN IGV y se validarán tasa parametrizada, precisión y redondeo antes de aprobar las plantillas.

Reportes principales en este orden: presupuestos ganados, OC completadas, OC con acta y OC pendientes de factura. Son vistas superpuestas: sus cantidades no se suman como si fueran servicios diferentes. Separar importes por moneda y contar una sola revisión comercial aceptada.

## 6. Arquitectura propuesta

**Decisión definitiva del usuario, 03/11/2026: ASP.NET Core MVC .NET 10, C#, ADO.NET y procedimientos almacenados SQL Server**, siguiendo la tecnología de FralseCont. Vistas .cshtml, ViewModels y controladores separados, con parciales reutilizables para grillas, filtros, fichas, popups, calendario y auditoría. No hay Blazor ni Razor Pages en la solución activa.

Los controladores coordinan solicitudes y validación; las reglas de negocio permanecen en Application/Domain. Infrastructure accede a SQL exclusivamente mediante ADO.NET y SP. Las vistas no consultan directamente la base. Los scripts se entregan para ejecución manual; no hay EF ni migraciones automáticas.

La futura app móvil podrá reutilizar reglas y servicios mediante una API autenticada; las vistas MVC no se transforman automáticamente en pantallas MAUI. Se prioriza la web MVC solicitada; la alternativa anterior de componentes Blazor queda sustituida por esta decisión.

Acceso: ASP.NET Core Identity o una integración equivalente con su gestión segura de contraseñas; concretar su almacenamiento antes de implementar. Las consultas y operaciones del negocio conservarán ADO.NET + SP. No guardar contraseñas en texto plano ni desarrollar cifrado propio.

Infraestructura de referencia: `DOCUMENTACION_INFRAESTRUCTURA_FRALSE.md` de la raíz del repositorio, aplicable por estar en `.NET2026`. Definir un servicio y base independientes; dominio y puerto pendientes. No reutilizar nombres de contenedores, conexiones ni secretos de FralseCont. Adjuntos privados fuera de wwwroot, acceso autorizado y respaldo coordinado con la base.

## 7. Plan por entregables

| Etapa | Trabajo | Criterio para terminar |
|---|---|---|
| 1. Base técnica y SQL | Solución por responsabilidades, tablas, catálogos internos, claves, seguridad y auditoría base; paquete SQL ordenado | Compilación y validación offline verificadas, diccionario y scripts entregados; sin ejecución SQL por Codex |
| 2. Acceso y presentación | Login, permisos en host, menú y componentes comunes con identidad HYT | Usuario autorizado entra, desactivado no accede; grillas, fichas, popups y fechas uniformes |
| 3. Mantenimientos | Clientes, tiendas, contactos, portal documental, unidades, zonas y materiales | Búsquedas, ayudas y paginación SQL de 20 con auditoría |
| 4. Tarifarios | Versiones, PEN/USD opcional, zonas, tipo y tasa IGV | Cambiar tarifa no altera documentos históricos; precio USD ausente no se convierte en cero |
| 5. Presupuestos | Una tienda, Invitación/Emergencia, revisiones, cálculos y dos formatos HTML | Recorrido comercial probado y formatos anterior/nuevo conciliados |
| 6. Seguimiento documental | Una OC y una guía por presupuesto, archivos y confirmación del cliente | Emergencia sin OC previa; archivos consultables, aceptación y auditoría visibles |
| 7. Marcado Facturado | Fecha, tipo 01/03, serie y número externos; sin emisión | Bloqueo sin requisitos y guardado con auditoría; reversión autorizada con motivo |
| 8. Dashboard y reportes | Indicadores, curvas interactivas, columnas, dona, vistas rápidas e impresiones HTML | Cifras conciliadas con listados, fechas identificadas y totales por moneda |
| 9. Validación y entrega | Recorridos, permisos, concurrencia, archivos, móvil, respaldo y despliegue | Aceptación funcional y pruebas en ambiente habilitado por el usuario |
| Posterior: Android | API autenticada y host móvil reutilizando servicios y reglas de negocio | Caso de uso compartido verificado en navegador y APK |

Como parte de la etapa 2 se construirán los componentes comunes de listado, ficha, trazabilidad e historial con identidad HYT. Cada etapa funcional deberá cumplir esos patrones. La aceptación incluirá comprobar alta y actualización, una transición válida y otra rechazada, historial de eventos repetidos, permisos de consulta, conservación de filtros y visualización en celular. Se verificará que un fallo al guardar la auditoría no deje aplicado el cambio de estado.

No se fijan plazos sin resolver el alcance pendiente. La primera entrega funcional será acceso, menú, configuración única y clientes con tiendas/contactos. Después se desarrolla un recorrido completo antes de ampliar los reportes.

## 8. Decisiones pendientes

1. Resuelto: una tienda específica por presupuesto, una guía por presupuesto y relación 1:1 entre OC y presupuesto (OC opcional hasta recibirla).
2. ¿Existen trabajos, actas y facturas parciales? ¿Una factura agrupa varias OC?
3. Resuelto: tarifario y presupuesto seleccionan CON IGV o SIN IGV. Confirmar tasa parametrizada, precisión y redondeo; no fijar impuestos a partir de ejemplos.
4. Resuelto para esta etapa: solo marcado Facturado y referencia del comprobante externo. Emisión electrónica e integración tributaria quedan fuera del alcance actual.
5. ¿Qué evidencia acredita aceptación del acta y confirmación para facturar? ¿Quién puede registrarlas o revocarlas?
6. Confirmar roles, formatos de OC y acta, significado de códigos del presupuesto e información histórica a importar.

La pregunta de una o varias empresas del levantamiento queda resuelta: una sola empresa HYT. Clientes y tiendas no son empresas operadoras del sistema.

## Aplicación inicial de identidad visual

03/11/2026: se aplica en la base web el logo HYT original y un tema común rojo, blanco y negro. Blanco para superficies, negro para estructura y texto, rojo para acento y acciones principales. Estilos centralizados en wwwroot/hyt-theme.css con variables de color, componentes de tablas y botones, foco visible, adaptación móvil y reglas iniciales de impresión. El inicio informa que los módulos y persistencia todavía están pendientes; no presenta datos ficticios ni operaciones disponibles. Los archivos CSS y recursos en wwwroot se incluyen mediante el SDK web. Compilación y revisión visual se registrarán según resultados.

## Decisiones consolidadas: IGV, OC, dólares y ejecución SQL

Confirmadas por el usuario el 03/11/2026. Prevalecen sobre propuestas y preguntas anteriores:

- **IGV:** crear catálogo TipoIGV con CON IGV y SIN IGV. El tarifario tendrá tratamiento predeterminado y cada partida conservará el suyo. En el presupuesto se elegirá el tratamiento de los precios y se convertirán al incorporar partidas según corresponda. Cambiar el selector recalcula sin aplicar el impuesto dos veces y conserva auditoría.
- **Conversión de precio:** siendo t la tasa parametrizada, SIN IGV → CON IGV: precio × (1+t); CON IGV → SIN IGV: precio ÷ (1+t). Mismo tratamiento: no convertir. SIN IGV significa que el precio no incluye el impuesto, no que la operación esté exonerada. Subtotal, impuesto y total se calculan coherentemente; tasa y criterio de redondeo se conservarán por revisión. No se fija aquí una tasa legal ni se añaden exoneraciones.
- **OC:** un presupuesto puede tener cero o una Orden de Compra - Cliente; no se permitirá asociarle una segunda. Cada OC corresponde a un solo presupuesto y cada presupuesto admite una sola guía y una sola tienda.
- **USD:** precio manual opcional por partida de tarifario, sin tipo de cambio automático. El campo vacío se conserva como no informado, nunca como cero. No bloquear un tarifario o presupuesto PEN por ausencia de USD. Si se elige USD y falta el precio de una partida, permitir guardar borrador e identificar el dato pendiente; no emitir una valoración final aparentemente completa. El comportamiento final de envío con precios faltantes se concretará antes de implementarlo.
- **SQL:** Codex prepara y entrega scripts; no crea la base ni ejecuta cambios contra servidores. Ubicaciones según AGENTS.md: Basededatos/Dbsishyt/Tablas, StoreProcedure y Script, relativas a la raíz del repositorio. El usuario ejecutará los scripts en el ambiente que corresponda. La conexión no bloquea la definición ni la generación de scripts.

Pruebas previstas: conversión en ambos sentidos y sin cambio, precios ya convertidos sin doble IGV, presupuesto sin OC y bloqueo de una segunda, tarifa sin precio USD y revisión de datos faltantes al seleccionar USD. Los importes históricos no se recalculan al cambiar un tarifario.

## Cierre de relaciones y dashboard inicial

Confirmado el 03/11/2026: una tienda específica obligatoria por presupuesto; una sola guía por presupuesto; una OC corresponde a un solo presupuesto y no agrupa varios. El presupuesto admite cero o una OC y cero o una guía durante su avance. Validar que tienda, cliente, presupuesto, OC y guía correspondan entre sí; impedir duplicados también en SQL.

La tasa inicial de IGV será **0.18**, en tabla independiente con vigencias, no como constante del programa. TipoIGV sigue distinguiendo CON IGV/SIN IGV. Al cambiar la tasa se agrega una nueva vigencia sin reescribir el histórico; cada revisión conserva la tasa aplicada. El redondeo sigue por concretar. Referencia de la tasa general: https://orientacion.sunat.gob.pe/3053-concepto-tasa-y-operaciones-gravadas-igv-empresas .

### Dashboard de inicio: entrega funcional obligatoria

Panel profesional de acceso rápido, con colores HYT y datos reales, adaptable a escritorio/tableta/celular. Filtros de periodo (último mes inicialmente), cliente y tienda; importes separados por moneda. Identificar la fecha usada en cada gráfico. Mostrar fecha/hora de actualización, estados de carga, sin datos y error. Evitar gráficos 3D y datos decorativos.

| Vista | Contenido y acción |
|---|---|
| Indicadores superiores | Cantidad de presupuestos, OC registradas, OC sin confirmación del cliente, confirmadas pendientes de marcar Facturado y marcadas Facturado |
| Líneas curvas interactivas | Evolución de presupuestos y OC registrados por día/semana, con curvas suaves, puntos y detalle de cada periodo; fecha de registro identificada, sin sumar sus importes como ingresos distintos |
| Barras/columnas (torres) | Cantidad de pendientes por cliente o etapa; separar las series y permitir abrir el listado filtrado |
| Torta o dona | Distribución de OC en categorías excluyentes: sin confirmación, confirmadas pendientes de marcado y marcadas Facturado; cantidades, porcentajes y total claros |
| Bandeja de acción | Presupuestos sin OC, sin guía, emergencias por regularizar, OC sin confirmar y por marcar Facturado; antigüedad y acceso directo a la ficha |

Las tarjetas abren los listados con sus filtros. Los pendientes de guía y OC pueden solaparse: no sumarlos como expedientes únicos. Para pendientes se propone mostrar tanto los del periodo como acceso al total pendiente a la fecha, evitando ocultar atrasos antiguos; rotular ambos alcances. «Confirmado» significa confirmación interna del cliente registrada en la OC, no aprobación del presupuesto. «Facturado» solo es el marcado manual con referencia externa.

Incorporar el dashboard en la etapa de consultas sobre las operaciones ya implementadas, con estructura visual desde la base inicial. Aceptación: cifras conciliadas con listados, categorías sin duplicación, totales por moneda, filtros coherentes y navegación rápida. No presentar una tasa de conversión comercial usando cohortes temporales distintas sin explicarlo.

### Curvas de movimiento e interacción del dashboard

Preferencia confirmada el 03/11/2026: gráfico de evolución con líneas curvas suaves y movimiento discreto al aparecer o cambiar filtros. Mostrar puntos reales y detalle al pasar el cursor, tocar o enfocar por teclado: fecha/periodo, serie y cantidad. Leyenda para mostrar/ocultar series, alternancia día/semana/mes y acceso al listado correspondiente al seleccionar un punto. Presupuestos y OC como series principales; eventos de confirmación y marcado Facturado como series seleccionables, con fecha y significado identificados.

Las curvas son una presentación de datos reales, no una simulación de actividad. No inventar puntos ni mostrar animación continua sin cambios de datos. La interpolación no debe introducir valores negativos o picos ajenos a los registros; periodos incompletos o sin datos se identifican. Mostrar tabla/resumen accesible alternativo. Respetar preferencias de movimiento reducido y uso táctil. Diferenciar series por etiquetas, trazos y puntos además de rojo/negro, con rellenos suaves sobre blanco. Contar eventos de cambio de estado sin duplicar correcciones o reenvíos; las reglas de conteo se conciliarán con auditoría al implementar.

## Estado de etapa 1

Preparación de archivos concluida el 03/11/2026 según ESTADO_ETAPA_1.md: solución compilada, 46 tablas, cuatro procedimientos base, cargas iniciales, protecciones y paquete de 59 scripts validados offline. No se ha conectado ni ejecutado SQL Server; su validación de despliegue queda a cargo del usuario en ambiente de pruebas. El acceso funcional corresponde a etapa 2.

## Estado de etapa 2

03/11/2026: acceso, políticas de permisos, cierre de sesión y componentes comunes implementados y verificados localmente. Paquete incremental de seis scripts preparado para ejecución manual posterior; sin conexión SQL realizada. La recuperación de acceso es asistida por el administrador mediante herramienta segura de restablecimiento. Véanse ESTADO_ETAPA_2.md y GUIA_ACCESO_ETAPA_2.md. Las pantallas CRUD de negocio y el dashboard conectado se mantienen en las etapas siguientes.

## Ajuste definitivo a MVC

03/11/2026: acceso, cierre, inicio y módulos ahora son controladores MVC y vistas .cshtml; las ayudas, grillas, filtros, fichas y trazabilidad usan parciales MVC. Calendario y popups funcionan con JavaScript del host. Se retiraron proyecto UI, Razor Components, Razor Pages y pruebas bUnit de la solución activa. ADO.NET, SP, seguridad de cookies y SQL se conservan. Ver AJUSTE_MVC.md.

## Estado de etapas 3 y 4

03/11/2026: mantenimientos, ayudas paginadas, auditoría visible, tarifarios versionados, publicación/copia y configuración de vigencias IGV implementados y verificados localmente en MVC/ADO.NET/SP. Paquetes incrementales preparados para ejecución manual; sin conexión a SQL Server. Véanse ESTADO_ETAPAS_3_Y_4.md y GUIA_ETAPAS_3_Y_4.md para cobertura y límite de la verificación. La etapa siguiente es presupuestos (5); no se ha importado automáticamente el Excel ni construido aún el cálculo o impresiones de presupuesto.

## Decisión vigente: Cloudflare R2 y Turnstile — 03/11/2026

Por instrucción del usuario, los adjuntos de la aplicación se guardan en Cloudflare R2. El administrador creará luego el bucket `hyt-documentos`, declarado en `.env.example` como CloudflareR2__BucketName y en el valor predeterminado de CloudflareR2Settings. Esta decisión reemplaza el almacenamiento local descrito en avances anteriores; el servicio local se conserva solo como sustitución en pruebas aisladas. No se creó ningún bucket ni se guardaron claves reales.

R2DocumentFileStore es el proveedor registrado en DI. CloudflareR2ObjectClient utiliza SDK S3, endpoint derivado de AccountId, credenciales explícitas, región auto, HTTPS, escritura condicional If-None-Match y opciones de compatibilidad R2. Validación común DocumentUploadRules conserva restricciones y SHA256. Reintento no sobrescribe contenido, lectura verifica integridad y permisos SQL/MVC siguen vigentes. No hay fallback a almacenamiento local ni enlaces públicos. No se requiere cambio de tablas o SP para el proveedor: la clave de objeto conserva el formato GUID/extensión ya validado por SQL.

Se integró Turnstile en login: widget público con acción login y validación Siteverify del lado servidor antes de acceder al repositorio de usuarios. Verifica éxito, acción, hostname permitido y vigencia. Falla cerrada ante token ausente, inválido o error de red. En producción es obligatorio; las pruebas lo desactivan explícitamente en Development o simulan el servicio. Claves exclusivamente por entorno, secret key ausente del HTML. Configuración y preparación en CONFIGURACION_CLOUDFLARE.md; variables sin valores secretos en .env.example. Los archivos .env privados están ignorados.

Pruebas nuevas de R2: conservación de contenido, huella, reintento sin duplicación/sobrescritura y rechazo de objetos alterados. Turnstile: éxito, acción y dominio ajenos, expiración, token ausente/largo, rechazo del servicio, JSON inválido, red y obligatoriedad en producción. Pruebas HTTP de widget sin secreto y bloqueo antes de consultar usuario al rechazar un token; éxito de Siteverify permite continuar a contraseña. No son pruebas contra un bucket real, pendiente de creación/configuración por el usuario conforme a su instrucción.

Avance previo de este tramo: regularización comercial de emergencia con OC aplicada en SQL (solo enviar/aprobar revisión vinculada, sin editar contenido/importes); historial consolidado del expediente paginado y filtrado por permisos; última actualización documental visible en listado de presupuestos. 90 comprobaciones SQL reales aprobadas, incluidas reversión completa ante fallo de auditoría, regularización sin recalcular, historial repetido de más de veinte eventos y protección de documentos sin permiso. OC muestra moneda/importe/confirmación y filtro con/sin confirmación. Continúan pendientes conciliación final de etapas 5/6, pruebas de navegador de los últimos ajustes y paquetes finales.
## Cierre de implementación y verificación de etapas 5/6 — 03/11/2026

Presupuestos y seguimiento documental conciliados contra el plan en VERIFICACION_ETAPAS_5_Y_6.md; recorridos en GUIA_ETAPAS_5_Y_6.md. Paquetes ETAPA_5_HYT.sql y ETAPA_6_HYT.sql y SHA256_ETAPA_5/6.json preparados en entregables. Instalación desde paquetes aprobada con 97 comprobaciones SQL reales; 164 comprobaciones HTTP/C# aprobadas. Chrome verificó presupuestos, seguimiento, ayudas, archivos, estados, regularización, filtro de OC confirmadas y presentación móvil/escritorio. Siete casos de impresión aprobados y contenido dentro de A4, con encabezados repetidos y todos los renglones conservados.

Validación offline: 142 fuentes, 47 tablas y 105 relaciones, más cinco pruebas de regresión del validador. Base local: 47 tablas, 78 SP y cero presupuestos/adjuntos de prueba. Los 28 SP de etapas 5/6 se compararon con sus definiciones instaladas en SistemaWebHYT y coinciden con las fuentes. Compilación sin advertencias ni errores; host y credenciales ficticias de prueba cerrados/eliminados.

Presupuesto_Obtener agrega EstadoOC y RevisionOC para regularización y consolida historial de presupuesto/revisión seleccionada, OC, guía, ejecución, confirmaciones y archivos, limitado por permisos específicos. Presupuesto_Listar toma actualización y responsable del presupuesto raíz, por lo que refleja cambios documentales. Se conserva paginación de veinte y conteo coherente. UI identifica la entidad con etiqueta legible y ofrece navegación al expediente completo desde popups.

User Secrets propio ya contiene la conexión SQL local, con autenticación de Windows, y campos Cloudflare. R2 y Turnstile están implementados y probados con sustitutos; el usuario creará el bucket hyt-documentos y completará las claves después, según su instrucción. La prueba real contra Cloudflare no se declara realizada. Aprovisionar el primer administrador según la guía de acceso; no hay credenciales predeterminadas. Marcado Facturado, dashboard y publicación permanecen en etapas 7/8/9.

## Entrega técnica de etapas 7/8 y preparación de 9 — 05/11/2026

Marcado Facturado y dashboard/reportes implementados y verificados localmente. El panel incluye cliente/tienda por nombre histórico, día/semana/mes, primera confirmación/primer marcado sin duplicar correcciones, navegación con filtros y bandeja paginada de pendientes del período/acumulados. Cuatro reportes HTML con totales por moneda y estilos de impresión. Evidencia y límites en VERIFICACION_ETAPAS_7_8_Y_9.md.

Decisión vigente del usuario: dejar preparada la publicación, sin activar todavía porque faltan claves y pruebas manuales. La etapa 9 entrega configuración y guía de aceptación, respaldos y reversión. No se declara aceptación en producción: faltan bucket/claves, dominio/parámetros definitivos, construcción/ejecución Docker y pruebas manuales/restauración en ambiente habilitado. No se realizaron cambios al servidor.
