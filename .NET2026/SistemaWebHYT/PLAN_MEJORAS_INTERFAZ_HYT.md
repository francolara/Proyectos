# Plan consolidado de mejoras de interfaz HYT

Fecha: 05/11/2026. Autor: FRANCO LARA.
Estado: implementado y verificado en entrega local. Evidencias consolidadas en VERIFICACION_INTERFAZ_HYT.md.

## Objetivo y alcance

Uniformar las pantallas de HYT a partir de las cuatro referencias del usuario. Mantener rojo, blanco y negro, la tipografía actual y la distribución de los formularios. Aplicar los cambios a los mantenimientos, presupuestos, seguimiento de OC/guías, marcado de facturado, tarifarios, tasas y consultas/reportes en pantalla que compartan los componentes afectados.

«Pantalla completa» significa ocupar todo el ancho disponible a la derecha del menú lateral, respetando márgenes y navegación. Los formularios de nuevo registro, edición y consulta individual conservan su ancho de lectura actual. Las grillas dentro de un formulario o popup conservan el ancho de su contenedor.

## Requisitos consolidados

| Mejora | Resultado esperado |
|---|---|
| Acciones de los listados | Ver, editar, imprimir, descargar y demás acciones existentes se presentan como botones; sin apariencia de enlace subrayado. Aplicar también a las tablas auxiliares y a eliminar cuando esa acción ya exista. |
| Historial en las ediciones | Retirar la sección general «Historial de cambios/acciones» de los formularios de registro y edición, incluidos los popups de edición. La auditoría interna se conserva. |
| Fechas y responsables | Franja compacta con título e icono, fondo suave y borde discreto de la paleta HYT. En escritorio, cuatro columnas: usuario de registro, fecha de registro, usuario de actualización y fecha de actualización. |
| Popups | Cabecera destacada con barra roja HYT, título blanco de mayor peso, explicación breve cuando corresponda y cierre visible. Cuerpo blanco, contenido desplazable y acciones claramente separadas. |
| Títulos de las pantallas | Encabezado definido sobre superficie blanca con acento rojo, título negro y subtítulo secundario. Botones principales alineados con el encabezado. |
| Ancho de los listados | Listados principales al ancho disponible, sin la limitación central que hoy deja grandes espacios laterales. Filtros y grilla alineados. |
| Modo oscuro | Selector de modo claro/oscuro en la barra del menú, disponible también en móvil y con preferencia recordada en el navegador. Adaptar fondos, textos, grillas, formularios, trazabilidad y popups manteniendo el acento rojo HYT y contraste legible. |
| Descripción de materiales y servicios | Área de texto que ocupa una fila completa y muestra inicialmente tres líneas; crecimiento vertical permitido. Conserva su contenido, obligatoriedad y límite actual de 1500 caracteres. |

## Etapa 1. Componentes visuales comunes

1. Definir las variantes de botón de acción: principal rojo para altas/guardados; secundario blanco y negro para consulta; edición con acento HYT; acciones destructivas con su confirmación existente.
2. Corregir los estilos compartidos para eliminar el subrayado de las acciones y evitar que las reglas actuales de grilla anulen el espaciado de los botones.
3. Crear un encabezado de pantalla reutilizable y una variante de contenedor ancho para los listados. Declarar el ancho por tipo de vista, evitando que una búsqueda, paginación o retorno desde edición cambie la distribución.
4. Conservar el código interno en el encabezado, debajo de «Volver al listado» y fuera del formulario.
5. Incorporar un selector claro/oscuro en la barra del menú. Centralizar los colores de ambas variantes para que todos los componentes cambien de forma uniforme. Recordar la elección en el navegador; en la primera visita usar la preferencia del sistema. Evitar destellos del tema incorrecto al cargar y mantener el selector accesible en el menú móvil.

Base de trabajo: hyt-theme.css, hyt-components.css, Views/Shared/_Layout.cshtml, Views/Shared/_Grid.cshtml y los modelos de presentación que requieran declarar el ancho.

## Etapa 2. Trazabilidad compacta y retiro de historial en edición

1. Unificar «Fechas y responsables» con el componente de trazabilidad, reemplazando las cuatro líneas verticales por la franja de cuatro columnas de la referencia.
2. Mostrar el usuario disponible y la fecha con el formato Perú existente. Cuando no haya actualización: «Sin actualizaciones». En registros nuevos aún no guardados, no inventar fechas ni responsables.
3. Adaptar a dos columnas en tablet y una columna en móviles estrechos, permitiendo nombres largos sin desbordamiento.
4. Retirar el historial general de cambios de las ediciones. Revisar tanto el historial del maestro como el que está incluido en el componente compartido y los formularios de operaciones.
5. Mantener la información funcional necesaria para trabajar con versiones, revisiones, adjuntos y estados documentales. Retirar la presentación del historial no elimina eventos ni sus controles de auditoría.

Base de trabajo: Views/Shared/_Trace.cshtml, Views/Master/_FormContent.cshtml y las vistas de presupuesto, seguimiento, empresa y marcado de facturado que muestran datos equivalentes. La consulta histórica de documentos se revisará por separado del formulario de edición para conservar las evidencias operativas existentes.

## Etapa 3. Cabeceras de todos los popups

1. Aplicar la cabecera HYT a ayudas de selección, zonas, calendario, tarifas contextuales y operaciones documentales: OC, guía, confirmación, ejecución y marcado como facturado.
2. Usar la referencia aportada para la jerarquía de título, cuerpo y pie; adaptar su verde y azul al rojo, blanco y negro de HYT.
3. Mantener un solo título principal por popup. En el editor de tarifarios embebido, coordinar el título exterior con el del contenido para evitar cabeceras repetidas.
4. Revisar cierre, Escape, retorno del foco, navegación por teclado y posición del indicador de carga. Mantener el bloqueo de doble envío y de cierre durante guardados.
5. Limitar la altura a la ventana disponible y desplazar el cuerpo del popup, conservando accesibles el título y sus acciones en pantallas pequeñas.

Base de trabajo: _Modal.cshtml, _Layout.cshtml, diálogos de Master y Budget, vistas CompactForm/Form y scripts hyt-ui.js, hyt-documents.js, hyt-dates.js y hyt-material-tariffs.js cuando sea necesario.

## Etapa 4. Aplicación a pantallas y materiales

1. Llevar al ancho completo los listados principales de maestros, presupuestos, seguimiento, tasas y reportes de consulta en pantalla.
2. Actualizar las acciones de _Grid y las tablas que tienen marcado propio: precios del material, adjuntos, revisiones, reportes y bandejas del dashboard.
3. Revisar todas las cabeceras de listado/registro para aplicar el mismo tratamiento visual y conservar filtros, botones y código interno en su lugar.
4. Cambiar únicamente la descripción de Materiales y servicios a un área de texto de tres líneas, con ancho completo de la grilla del formulario. Mantener SAP, unidad, especialidad y demás campos con su distribución actual.
5. Conservar las dos impresiones HTML de presupuestos y sus estilos de impresión; los encabezados visuales de navegación no deben agregarse al documento impreso.

Base de trabajo: vistas Index/Form/Details de los módulos afectados, _MaterialTariffs.cshtml, _Attachments.cshtml, metadatos de Materiales en maintenance_schema.py/Definitions.cs y el renderizador compartido del formulario.

## Etapa 5. Validación y entrega local

- Compilar la solución MVC y ejecutar las comprobaciones existentes afectadas por cambios de controles, cabeceras y visibilidad del historial.
- Revisar en navegador escritorio, tablet y móvil: maestros, presupuestos, OC/guías, tarifas, reportes y ayudas de selección.
- Comprobar que las tablas aprovechan el ancho sin recortar acciones; en móvil, el desplazamiento horizontal queda dentro de la grilla.
- Probar búsquedas, paginación, guardados, errores, carga, permisos y retorno desde edición. Conservar la paginación SQL de veinte registros.
- Confirmar el formato de fechas, la franja de responsables y la ausencia del historial en edición; verificar que se siguen registrando eventos de auditoría.
- Verificar que el texto de material admite saltos de línea y conserva su valor al guardar, editar y consultar. Revisar su representación en presupuesto/impresión si reutiliza esa descripción.
- Revisar capturas de antes/después y los dos formatos de impresión HTML para detectar efectos de estilos compartidos.
- Probar ambos temas en todas las pantallas y popups: contraste, foco, campos deshabilitados, errores, gráficos e indicador de carga. Comprobar persistencia al recargar, funcionamiento en móvil y sincronización de las ventanas de tarifas embebidas. Las impresiones HTML conservan fondo blanco y su diseño documental.
- Actualizar firmas del día en formularios modificados, comprobar UTF-8/CRLF y revisar el alcance de los cambios.
- Actualizar la compilación normal que utiliza Visual Studio, resolviendo el bloqueo de la ejecución local si está activa, para evitar mostrar una versión anterior. Entregar para revisión con F5, sin despliegue remoto.

## Criterios de cierre

1. Ninguna acción de fila existente mantiene apariencia de enlace simple.
2. Los listados principales ocupan el ancho disponible; los formularios conservan su ancho actual.
3. Títulos de pantalla y popup se distinguen del fondo y mantienen identidad HYT.
4. Las ediciones muestran una franja compacta de fechas/responsables y no presentan el historial general de cambios.
5. La descripción del material ocupa toda la fila y muestra tres líneas inicialmente.
6. Permisos, filtros, carga, adjuntos, auditoría interna y formatos de impresión siguen funcionando.
7. El selector claro/oscuro está visible en la barra del menú y recuerda la preferencia; las pantallas y popups mantienen legibilidad e identidad HYT en ambos modos.

Implementación realizada en vistas, estilos, scripts y componentes MVC. Se conservan la estructura SQL y los procedimientos existentes.

## Corrección vigente — 05/11/2026

El historial general de cambios/acciones queda oculto en todos los lugares donde se presentaba, incluidas las consultas de maestros, presupuestos y seguimiento y las ventanas embebidas. Este alcance sustituye las menciones anteriores a conservar su presentación en consulta. Se mantienen Fechas y responsables, las revisiones/adjuntos operativos y la auditoría interna. Pruebas HTTP de consulta en seis rutas aprobadas; no cambia SQL.

Actualización 05/11/2026: la ventana de tarifas del material se sustituye por **Agregar tarifario**, con IGV, PEN, USD opcional y estado; se retiran los pasos de versiones/publicación del recorrido del material. Alta y edición aprobadas en navegador; regresión de interfaz: 255 comprobaciones aprobadas. Ver CODIGOS_Y_TARIFAS_DESDE_MATERIALES.md.
