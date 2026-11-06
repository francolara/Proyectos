# Revisión de presupuestos Excel para reportes HTML

Fecha: 03/11/2026. Revisión de contenido, estructura, fórmulas y configuración de impresión del paquete XLSX; no se realizó una comparación visual renderizada ni se modificaron los originales.

## Fuentes revisadas

- `C:/Users/Franco Lara/Downloads/H Y T - PPTO 100(ANT).xlsx`.
- `C:/Users/Franco Lara/Downloads/H Y T - PPTO PS0 (NUEVO).xlsx`.

## Viabilidad

Ambos pueden reproducirse como reportes HTML con tablas, cabecera corporativa, datos comerciales, partidas, totales, consideraciones y firmante. No requieren Excel para mostrarse una vez implementados los datos y cálculos en el servidor. La fidelidad visual exacta y la paginación se validarán con una muestra HTML e impresión reales; todavía no existe una plantilla HTML implementada.

| Aspecto | Antiguo: Hoja1 | Nuevo: Hoja1 |
|---|---|---|
| Cabecera | Identificación HYT, RUC, servicios y contacto | Identificación HYT, RUC, servicios y contacto |
| Datos | Descripción, cotización, fecha, tienda, coordinador/encargado y código de proveedor | Cliente, dirección, presupuesto, fecha, trimestre, zona, vigencia, proveedor SAP, contrato marco, especialidad, cotización y ejecución |
| Detalle | Ítem, descripción, unidad, cantidad, precio unitario, parcial y total de grupo | Tienda, centro SAP FICO, necesidad, tipo de servicio, posición, material SAP, descripción, unidad, cantidad, precio PEN y subtotal |
| Totales | Costo final sin IGV | Subtotal, IGV 18 % y total |
| Cierre | Notas, saludo y firmante/cargo | Consideraciones, saludo y firmante/cargo |
| Impresión configurada | A4 vertical | A4 vertical; tabla de 11 columnas que requiere prueba de legibilidad |

## Hallazgos que afectan al sistema

1. Ambos libros contienen referencias a un libro externo para recuperar datos. La web sustituirá esas búsquedas por datos del maestro de clientes, tiendas y contactos, sin depender de rutas de Excel.
2. En el nuevo hay resultados guardados `#N/A` en referencias a tienda y centro SAP. Son datos no resueltos del archivo, no campos que deban imprimirse así; la aplicación validará su disponibilidad.
3. El antiguo usa un total por grupo y parciales; no se debe sumar nuevamente ambos como si fueran partidas independientes. El nuevo también tiene una fila de agrupación en K19 que suma K20:K30.
4. En el nuevo, J32 contiene 174, J33 calcula 31.32 y J34 suma 205.32. Sin embargo, A36 dice «Los costos incluyen IGV». Confirmar si los precios unitarios incluyen o excluyen impuesto antes de fijar cálculos y textos. La tasa del ejemplo no se adopta automáticamente como parámetro universal.
5. La Hoja2 oculta contiene referencias de tiendas y responsables. No es un reporte adicional; sirve como fuente para evaluar una futura carga de maestros, previa depuración.
6. El nuevo tiene una Hoja3 con valores de especialidad en H1048567:H1048568. No debe convertirse en millones de filas ni imprimirse como parte del presupuesto.
7. El nuevo contiene «PPTO N°» y «N° COTIZACION»: confirmar si identifican documentos distintos o referencias del cliente; no unificarlos sin validación.

## Diseño recomendado

Las pantallas de captura mantienen un formato común, con secciones de cabecera, partidas y resumen. Alcance confirmado por el usuario el 03/11/2026: la impresión incluirá dos variantes HTML seleccionables, **Formato nuevo** y **Formato anterior**; ambas son entregables obligatorios del plan de trabajo. La opción predeterminada queda por definir. Elegir una variante no cambia los datos ni las reglas de cálculo del presupuesto.

Para pantalla: tabla adaptable con desplazamiento horizontal cuando sea necesario, sin ocultar códigos importantes. Para impresión: ocultar menú y botones, repetir encabezado de columnas, controlar saltos de página y mantener totales/firmante legibles. Probar primero A4 vertical como las referencias; ofrecer horizontal si resulta necesario y se valida el cambio. No reducir toda la tipografía para forzar una página.

Los reportes operativos también serán HTML: título, fecha de generación, filtros aplicados, tabla de resultados y totales por moneda. Los presupuestos conservarán su composición documental. Uniformidad significa compartir estilos y controles, manteniendo los campos necesarios de cada documento.

## Criterios de aceptación de las futuras plantillas

- Comparar campos y orden con cada referencia; validar logo, anchos, importes y saltos de página.
- Probar cero partidas, una partida, múltiples páginas y descripciones largas.
- Repetir encabezados del detalle y evitar cortes en totales y firmas.
- Calcular importes con precisión decimal en servidor y conciliar con ejemplos aprobados.
- No mostrar errores de Excel ni depender del libro externo.
- Conservar datos, precios y versión del formato utilizado por cada revisión enviada.
- Aplicar permisos a consulta e impresión y registrar la generación en auditoría.

No se implementaron pantallas, importación de maestros ni cambios de estructura SQL en esta revisión.

## Implementación posterior — 03/11/2026

Ambas variantes HTML están implementadas en Views/BudgetReports, con CSS autónomo y generación conservada/auditada. Se revisaron otra vez valores y combinación de celdas de Hoja1 de ambos archivos originales, en modo de solo lectura. Los textos ambiguos sobre IGV se sustituyen por el tratamiento real de la revisión; moneda y tasa se muestran desde sus snapshots. No se copian errores #N/A ni vínculos externos.

Chrome/PDF: una partida en ambos formatos y borrador vacío/moneda USD caben en una página A4. Se preservaron cuarenta partidas en varias páginas, encabezados repetidos, total 4,720.00 y texto extenso completo. Capturas de páginas revisadas. La cabecera del formato nuevo se distribuye en dos pares de columnas y usa abreviaturas para evitar palabras partidas en las once columnas. No se modificaron los originales ni se requiere Excel para generar/imprimir.
