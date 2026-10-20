# Documentación inicial: núcleo de documentos de venta

**Proyecto:** `Ventas PROVEEPERU` (VB6)  
**Fecha de levantamiento:** 13/10/2026  
**Alcance inicial:** levantamiento técnico y funcional.
**Última actualización:** 19/10/2026 — implementación del control universal de importaciones por documento e ítem origen.

## 1. Resumen ejecutivo

El centro operativo de ventas está concentrado en tres formularios muy similares: `frmDocVentas`, `frmDocVentas_Mostrar_V` y `frmDocVentas_Vista`. Cada uno supera las 20 000 líneas. Comparten el ciclo completo de venta: listado, creación, edición, validación, cálculo, grabación, anulación, eliminación, impresión, electrónica, stock, caja, crédito y contabilidad.

La principal causa técnica de lentitud al usar una base remota es la granularidad de acceso: se ejecutan consultas SQL directas y búsquedas unitarias desde eventos de interfaz. Se detectaron aproximadamente 431, 430 y 429 apariciones de sentencias SQL directas, respectivamente. La lógica también está triplicada, por lo que una corrección de rendimiento debe centralizarse para no introducir diferencias entre las tres variantes.

## 2. Puntos de entrada

`frmPrincipal.frm` abre `frmDocVentas` con el tipo de documento indicado:

| Menú Ventas | Código | Formulario | Proceso |
|---|---:|---|---|
| Factura | 01 | `frmDocVentas` | Facturación |
| Boleta de Venta | 03 | `frmDocVentas` | Boleta |
| Pedido | 40 | `frmDocVentas` | Pedido; puede importarse a factura |
| Guía de Remisión | 86 | `frmDocVentas` | Movimiento/guía |
| Nota de Crédito | 07 | `frmDocVentas` | Crédito y referencia |
| Ticket | 12 | `frmDocVentas` | Ticket |
| Nota de Venta | 56 | `frmDocVentas` | Venta interna |
| Separación de Mercadería | 91 | `frmDocVentas` | Reserva/separación |
| Cotización | 92 | `frmDocVentas` | Cotización |

El menú también contiene Nota de Débito (código 08), Guías Madres, Liquidaciones, Atribuciones y Atribuciones–Notas de Crédito. Deben considerarse en la segunda fase porque comparten datos de `DocVentas` y de series.

## 3. Formularios revisados

### `Formularios/frmDocVentas.frm`

Es la implementación operativa principal. Recibe `strTipoDoc` y, opcionalmente, datos para importar un documento. Sus puntos de control son:

| Rutina | Línea aprox. | Responsabilidad |
|---|---:|---|
| `Form_Load` | 5729 | Inicializa pantalla, parámetros, grillas, documento, visibilidad y listado. |
| `nuevo` | 7667 | Prepara un documento nuevo, numeración/serie y estado de edición. |
| `listaDocVentas` | 13362 | Carga la bandeja/listado de comprobantes. |
| `mostrarDocVentas` | 13392 | Recupera cabecera, detalle y referencias para edición/consulta. |
| `Grabar` | 5968 | Valida y persiste cabecera, detalle, referencias y efectos asociados. |
| `anularDoc` / `eliminar` | 15214 / 17309 | Anula o elimina, incluyendo reversos según el documento. |
| `calcularTotales` | 16013 | Recalcula impuestos, descuentos y totales. |
| `AsientoContableVentas` | 19050 | Genera/reversa efecto contable. |
| `ValidaStock_Y_Saldos` | 21233 | Valida stock y saldos del producto. |

### `Formularios/frmDocVentas_Mostrar_V.frm`

Es una variante de visualización/versión, con los mismos procesos base. Su `mostrarDocVentas` recibe además `stritm`, por lo que distingue un ítem o versión al recuperar el documento. Contiene 20 649 líneas y conserva casi toda la lógica de grabación, anulación, cálculo, crédito, stock y contabilidad. A diferencia de los otros formularios, oculta la importación de atenciones en los tramos revisados.

### `Formularios/frmDocVentas_Vista.frm`

Es otra variante de consulta/visualización de documentos, con la entrada pública adicional `strNumdocumento`. Mantiene el mismo ciclo de listado, carga, grabación, anulación y eliminación; usa `spu_Docventas_Listacab_Consulta` para el listado. Tiene 21 024 líneas. Por su duplicación casi total, debe tratarse como consumidor de una futura capa común y no como una implementación independiente.

## 4. Flujo funcional de un documento de venta

```text
Menú → frmDocVentas(strTipoDoc)
     → Form_Load: parámetros + configuración + listado
     → Nuevo o selección del listado
     → Cabecera: cliente, moneda, almacén, forma de pago, serie
     → Detalle: producto, unidad, precio, descuento, impuesto
     → Validaciones: formulario, cierre de mes, referencias, stock, crédito, permisos
     → Grabar: documento, detalle, referencias y temporales
     → Efectos: stock / caja-cuentas por cobrar / electrónica / contabilidad / log
     → Impresión, aprobación, anulación o eliminación
```

### 4.1 Carga inicial

`Form_Load` obtiene la descripción y dimensiones del documento con `spu_Documento`, aplica configuraciones visuales, carga el listado y habilita acciones. Lee parámetros que determinan comportamiento, tales como venta electrónica, importación entre empresas, stock por lote, cambio de IGV, aprobación automática e importación de atenciones.

La aplicación ya posee una carga central: `cargarParametrosSistema` en `Modulos/mdlFunciones.bas` (líneas 1892–1959), invocada en `frmLogin` y nuevamente en `frmPrincipal`. Esa rutina llena variables globales `STR_*`; debe ser la vía preferida antes de abrir documentos.

### 4.2 Creación y edición

El botón de nuevo ejecuta `nuevo`, inicializa temporales de lotes/liquidación, controla qué botones quedan visibles y prepara el formulario. La serie condiciona el máximo de registros, los permisos, el tipo de documento y el correlativo. Al seleccionar un registro, `mostrarDocVentas` carga cabecera, detalle y referencias; al importar, se usan variantes `spu_Docventas_Listacab_Imp` y `spu_Docventas_ListaDet_Imp`.

### 4.3 Grabación y validaciones

`Grabar` realiza, entre otras, estas validaciones:

- cierre de mes;
- controles obligatorios de cabecera y detalle (`validaFormSQL`);
- detalle no vacío y referencias documentales;
- centro de costo y unidad de negocio;
- saldo y línea de crédito;
- numeración de factura/electrónica;
- stock, lotes, liquidaciones y pedidos importados;
- permisos, aprobación y restricciones de edición.

Los procesos asociados incluyen cuentas por cobrar, movimientos de stock, factura entre empresas, asiento contable, log de ventas y electrónica. El alcance exacto depende de `strTipoDoc`, parámetros y configuración de serie.

## 5. Servicios, procedimientos y dependencias relevantes

Procedimientos encontrados en el núcleo revisado:

| Dominio | Procedimientos / servicio |
|---|---|
| Documento | `spu_Documento`, `spu_Docventas_Listacab`, `spu_Docventas_Listacab_Consulta`, `spu_Docventas_ListaDet`, `spu_Docventas_Mostrar_Docventas_cab`, `..._Det`, `..._Ref` |
| Importación | `spu_Docventas_Listacab_Imp`, `spu_Docventas_ListaDet_Imp`, `spu_ObtenerPedidosImportadosDeGuia`, `spu_RecalculaImportacionPedido` |
| Stock | `Spu_CalculaStockNuevo`, `Spu_ValidaStock_Y_Saldos` |
| Contabilidad | `Spu_AsientoContableVentas`, `Spu_AsientoContableVentasTransGratuita`, `spu_genCuentaDestino`, `Spu_GenCuentaDestino_2` |
| Cliente/precio | `spu_traDatos_Cliente`, `spu_Recupera_Direccion_Cliente`, `spu_Trae_Costo_Precio_Cotizacion`, `Spu_VerificaMorosos` |
| Operaciones auxiliares | `spu_GrabaVentas_Log`, `spu_EliminaTemporales`, `spu_Docventas_Reg_AtnDet`, `Spu_CalculaFacturacionPorHC` |

`mdlFunciones.bas` concentra conexión ADO, consultas reutilizables (`traerCampo`, `traerCampos`), parámetros (`leeParametro`), grillas y ejecución de formularios. La función `leeParametro` hace una consulta por petición; en los formularios se mantiene un uso mixto de variables globales `STR_*` y llamadas repetidas a `leeParametro`/`traerCampo`.

## 6. Inicio y mantenimiento de series

### Series por sucursal y documento: `frmSeriesDocumento`

| Rutina | Responsabilidad |
|---|---|
| `Form_Load` | Configura grilla y carga las series de la sucursal. |
| `mostrarSeries` | Consulta `seriesdocumento` junto con `documentos`. |
| `Grabar` | Inicia transacción, elimina las series de empresa/sucursal e inserta lo mostrado en la grilla. |

Guarda: `idSucursal`, `idDocumento`, `idSerie`, `item`, `idEmpresa`, `numRegMaximo` y `espacioLineasImp`. Es crítico: una omisión o duplicación puede impedir el inicio de documentos o alterar sus límites de detalle/impresión.

### Series por usuario: `frmSeriesUsuario`

Permite autorizar qué series puede utilizar cada usuario. `mostrarSeries` consulta `seriexusuario`; `Grabar` reemplaza las asignaciones del usuario dentro de una transacción e inserta `idUsuario`, `idDocumento`, `idSerie`, `item` e `idEmpresa`. Ambos mantenimientos usan SQL concatenado y el patrón borrar–reinserir.

## 7. Catálogo de mantenimientos del menú

### Mantenimiento (`mnu02`)

| Opción | Finalidad funcional |
|---|---|
| Entidades | Maestro de entidades. |
| Clientes / Proveedores | Terceros comerciales usados en documentos. |
| Vendedores / Choferes / Vehículos / Emp. Transporte | Recursos comerciales y de despacho. |
| Productos / Marcas / Grupos de Productos / Tallas | Catálogo de artículos y clasificación. |
| Unidad de Medida / Niveles / Conceptos | Catálogos generales de clasificación. |
| Almacenes / Almacenes para Ventas | Disponibilidad y origen/destino de stock. |
| Listas de Precios / Formas de Pago | Valoración y cobranza de documentos. |
| Sucursales / Cajas / Tipos Mov. Caja | Operación por sede y caja. |
| Centros de Costo / Unidades de Producción | Imputación operativa y validación de detalle. |
| Número Máximo Registros | Límite de detalle asociado a serie/documento. |
| Zonas / Ubigeo / Motivos Traslados / Canales | Datos logísticos y comerciales. |
| Datos / Fórmulas | Datos auxiliares y fórmulas de negocio. |
| Documentos | Tipos documentarios, configuración base y serie predeterminada. |

### Configuración (`mnu08`), Seguridad (`mnu09`) y Utilitario (`mnu11`)

| Área | Opción | Relación con documentos de venta |
|---|---|---|
| Configuración | Conf. Documentos de Ventas | Define objetos, campos y comportamiento del formulario. |
| Configuración | Documentos | Define tipos documentarios disponibles. |
| Configuración | Series por Usuario | Abre `frmSeriesUsuario`; controla acceso a series. |
| Configuración | Conf. Impresión de Documentos/Motivos/Recibos/Etiquetas | Define salida impresa y etiquetas. |
| Configuración | Variables del Sistema | Fuente de parámetros `STR_*`. |
| Configuración | Tipos de Niveles / Tipos Mov. Caja / Documentos a Importar | Catálogos que afectan clasificación, caja e importación. |
| Seguridad | Usuarios / Perfiles / Claves de Autorización | Permisos de las acciones sensibles. |
| Seguridad | Cierre del Mes / Auditoría–Registro de Movimientos | Control de período y trazabilidad. |
| Utilitario | Series por Usuario | Mismo mantenimiento de asignación de series. |
| Utilitario | Cambio de Comisión, Caja o Vendedor; Liberar Documento | Cambios posteriores al documento. |
| Utilitario | Asientos Contables: generar, consulta, consistencia, transferir, eliminar | Reprocesamiento y control contable. |
| Utilitario | Procesar Tipos de Cambio / Transferir a Bancos / PDB | Procesos financieros complementarios. |
| Utilitario | Ventas Electrónicas: baja y resumen diario | Ciclo fiscal electrónico. |
| Utilitario | Importar Productos; consultas de entidades, clientes, proveedores y productos | Carga y consulta de maestros. |

## 8. Hallazgos y prioridad para nube

| Prioridad | Hallazgo | Impacto | Acción recomendada |
|---|---|---|---|
| P0 | SQL directo concatenado en formularios y mantenimientos de series. | Muchos viajes de red; riesgo de inyección, conversión y errores de comillas. | Sustituir por procedimientos almacenados parametrizados; empezar por listado, carga y series. |
| P0 | Consultas unitarias dentro de validaciones y eventos de grilla. | Latencia multiplicada por cada fila o tecla. | Crear operaciones por lote: cliente, precio, stock, permisos y parámetros necesarios para el documento. |
| P0 | Tres copias del flujo de `DocVentas`. | Corregir una variante no corrige las demás; alto riesgo funcional. | Extraer primero servicios comunes de datos y reglas sin alterar la interfaz VB6. |
| P1 | `leeParametro` y `traerCampo` siguen usándose en pantalla pese a la caché `STR_*`. | Consultas repetidas y comportamiento disperso. | Completar `cargarParametrosSistema`; proveer una caché tipada con refresco explícito al cambiar empresa/sucursal o configuración. |
| P1 | Series: borrar e insertar toda la asignación. | Bloqueos y riesgo ante concurrencia; más tráfico que el necesario. | SP transaccional con tabla de detalle/TVP y validación de duplicados. |
| P1 | Operación extensa y multiefecto en `Grabar`. | Difícil recuperar fallos parciales en red. | SP orquestador por documento, transacción explícita y resultado estructurado. |
| P2 | Temporales, electrónica, impresión y contabilidad ocurren alrededor de la misma pantalla. | Tiempo de respuesta percibido elevado. | Medir y separar tareas posgrabación que no necesiten bloquear al usuario. |

## 9. Ruta de intervención recomendada

1. Instrumentar tiempos y número de llamadas para abrir listado, nuevo, cargar documento y grabar cada tipo documental.
2. Completar la carga única de parámetros en sesión y reemplazar las lecturas restantes de parámetros en estos tres formularios.
3. Migrar a SP parametrizados las tres lecturas de mayor frecuencia: listado, cabecera/detalle/referencias y catálogos de cabecera.
4. Crear operaciones por lote para precios, stock y validaciones del detalle; no consultar por fila ni por pulsación.
5. Reemplazar el mantenimiento de series por dos SP transaccionales (`SeriesDocumento_Guardar` y `SeriesUsuario_Guardar`).
6. Extraer las reglas repetidas de los tres formularios hacia módulos/clases VB6 comunes antes de modificar reglas de negocio.
7. Migrar `Grabar`, anulación y eliminación por tipo de documento, con pruebas de regresión de stock, caja, crédito, contabilidad y electrónica.

## 10. Fuentes revisadas

- `Formularios/frmDocVentas.frm`
- `Formularios/frmDocVentas_Mostrar_V.frm`
- `Formularios/frmDocVentas_Vista.frm`
- `Formularios/frmSeriesDocumento.frm`
- `Formularios/frmSeriesUsuario.frm`
- `Formularios/frmPrincipal.frm`
- `Modulos/mdlFunciones.bas`
- `Modulos/mdlVariables.bas`
- `Ventas.vbp`

## 11. Control de importaciones por ítem — implementación del 19/10/2026

### 11.1 Problema corregido

La importación histórica identificaba las cantidades consumidas principalmente por producto y unidad de medida. Cuando un pedido contenía el mismo producto más de una vez, grabar una guía podía actualizar simultáneamente las dos líneas del pedido, aunque solamente se hubiera importado una de ellas.

Además, `ItemPro` se copiaba de manera transitiva. Por ejemplo, al crear un pedido desde una cotización y luego una guía desde ese pedido, la guía podía conservar el ítem de la cotización en lugar del ítem inmediato del pedido.

### 11.2 Regla implementada

`ItemPro` pasa a representar el `Item` del documento origen inmediato:

```text
Cotización, ítem 5
    → Pedido, ítem 2, ItemPro = 5
        → Guía, ítem 1, ItemPro = 2
            → Factura, ítem 3, ItemPro = 1
```

Esto permite que cada documento descuente exactamente la línea desde la cual fue importado, aunque el mismo código de producto aparezca varias veces.

La regla se aplica en la ayuda normal de importación, en la ayuda entre empresas cuando trabaja sin agrupación y en la generación `G. Pedido` desde una cotización.

### 11.3 Compatibilidad con documentos históricos

El procedimiento `dbo.spu_RecalculaImportacionDocumento` utiliza una estrategia mixta:

- Si un producto y unidad de medida aparecen una sola vez en el documento origen, acumula las importaciones por producto/UM. Esto incluye correctamente guías antiguas y nuevas.
- Si el producto/UM aparece en más de una línea, acumula por `ItemPro = Item` del origen. Las importaciones nuevas quedan separadas correctamente.
- Los documentos históricos con productos repetidos y relaciones antiguas ambiguas requieren corrección de datos; no se redistribuyen cantidades automáticamente porque no existe evidencia suficiente para identificar la línea correcta.
- Los documentos destino anulados se excluyen del acumulado.

Esta compatibilidad permite continuar atendiendo pedidos antiguos pendientes que no tengan productos repetidos sin perder las cantidades ya importadas.

### 11.4 Procesos cubiertos

El recálculo se ejecuta dentro de la transacción correspondiente al:

1. Grabar un documento nuevo.
2. Modificar un documento, considerando referencias retiradas y referencias nuevas.
3. Generar automáticamente un documento relacionado, incluido `G. Pedido` y la facturación automática desde guía.
4. Anular un documento.
5. Eliminar un documento.

El flujo anterior que incrementaba o descontaba `CantidadImp` directamente desde VB6 queda deshabilitado para `STR_AGRUPAPRODUCTOS = "N"`. En ese modo, el valor se reconstruye desde los documentos destino vigentes mediante el procedimiento almacenado.

El modo agrupado conserva su funcionamiento histórico porque, al combinar productos, deja de existir una correspondencia individual entre las líneas originales.

### 11.5 Procedimientos almacenados

| Procedimiento | Responsabilidad |
|---|---|
| `dbo.spu_Docventas_ListaDet_Imp` | Devuelve el detalle disponible y expone `Item AS ItemOrigen`; también cubre la carga automática de `G. Pedido`. |
| `dbo.spu_ObtenerDocumentosImportados` | Obtiene todos los documentos origen locales referenciados por un documento destino. |
| `dbo.spu_RecalculaImportacionDocumento` | Reconstruye `CantidadImp` y `EstDocImportado` del detalle y de la cabecera mediante la estrategia mixta. |
| `dbo.spu_RecalculaImportacionPedido` | Wrapper de compatibilidad para procesos existentes que todavía invocan el recálculo específico de pedidos. |
| `dbo.spu_Docventas_AuditaImportacionItems` | Auditoría de solo lectura que identifica importaciones históricas ambiguas cuando se repite producto/UM. |
| `dbo.spu_Docventas_ValidaEliminaItemImportado` | Informa qué líneas del documento origen ya fueron importadas y protege su identidad durante la edición. |

### 11.6 Archivos VB6 modificados

| Archivo | Cambio |
|---|---|
| `Formularios/frmListaDocExportar.frm` | Conserva el ítem real del detalle seleccionado en el campo temporal `ItemOrigen`. |
| `Formularios/frmListaDocExportarEntreEmpresas.frm` | Incorpora el mismo campo temporal para que la ruta sin agrupación sea compatible. |
| `Formularios/frmDocVentas.frm` | Guarda el ítem origen inmediato y centraliza los recálculos posteriores a grabar, modificar, anular y eliminar. |
| `Clases/clsDocVentas.cls` | Deshabilita el incremento directo por producto/UM en el modo sin agrupación y utiliza la variable global de configuración. |

No se modificó la estructura de `DocVentasDet`: se reutilizó `ItemPro`. Tampoco se modificaron archivos `.frx`.

### 11.7 Orden de despliegue

1. `dbo.spu_RecalculaImportacionDocumento.StoredProcedure.sql`.
2. `dbo.spu_ObtenerDocumentosImportados.StoredProcedure.sql`.
3. `dbo.spu_Docventas_ListaDet_Imp.StoredProcedure.sql`.
4. `dbo.spu_RecalculaImportacionPedido.StoredProcedure.sql`.
5. `dbo.spu_Docventas_AuditaImportacionItems.StoredProcedure.sql`.
6. `dbo.spu_Docventas_ValidaEliminaItemImportado.StoredProcedure.sql`.
7. Ejecutable VB6 recompilado con los formularios y la clase actualizados.

La auditoría puede ejecutarse para toda una empresa o limitarse progresivamente por sucursal, documento, serie y número. Devuelve una fila por cada posible ítem origen y detalle destino. `ITEM_ORIGEN_NO_COINCIDE` indica que el `ItemPro` histórico no corresponde con ese posible ítem; `REVISAR_ASIGNACION_HISTORICA` indica una coincidencia técnica que debe validarse funcionalmente porque el producto/UM está repetido.

### 11.8 Pruebas de regresión requeridas

- Pedido antiguo sin productos repetidos, parcialmente importado en una guía antigua; importar el saldo a una guía nueva.
- Pedido con dos líneas del mismo producto/UM; importar cada ítem en guías diferentes.
- Importar parcialmente una misma línea en más de un documento destino.
- Cadena cotización → pedido → guía → factura.
- Generar un pedido mediante `G. Pedido` y luego importarlo a una guía.
- Modificar cantidades o retirar una referencia de un documento ya grabado.
- Anular y eliminar un documento destino, comprobando la restitución del saldo del origen.
- Confirmar que un documento destino anulado no participe en `CantidadImp`.
- Confirmar que `EstDocImportado` sea `S` solamente cuando todas las líneas del origen estén completamente importadas.
- Intentar eliminar con Supr una línea origen ya importada y confirmar el bloqueo inmediato.
- Eliminar una línea no importada de un documento parcialmente importado y confirmar que los números `Item` restantes no cambien.
- Eliminar o insertar una línea en un documento nuevo o sin importaciones destino y confirmar que los ítems se renumeren correlativamente desde 1.
- Agregar una línea después de una eliminación y confirmar que recibe un número nuevo sin reutilizar ni duplicar un `Item` existente.
- Intentar grabar después de retirar o reemplazar indirectamente una línea importada y confirmar el bloqueo defensivo del guardado.

### 11.10 Protección de líneas origen importadas

Durante la modificación de un documento sin agrupación, `frmDocVentas` consulta `dbo.spu_Docventas_ValidaEliminaItemImportado` antes de borrar una línea. Si existe cantidad importada en un documento destino vigente, la eliminación se cancela y se informa el ítem, el tipo, la serie, el número y la cantidad correspondiente a cada documento destino.

La misma regla vuelve a comprobarse antes de grabar. Esta segunda validación evita que una línea importada pueda retirarse o sustituirse por una ruta distinta a la tecla Supr. La numeración se decide dinámicamente: los documentos nuevos y los documentos sin ninguna importación destino conservan el comportamiento anterior y se renumeran correlativamente; cuando existe al menos una línea importada, se preservan todos los números `Item`, aunque queden espacios, para no romper las referencias existentes.

El siguiente número de ítem se calcula en `gDetalle_OnBeforeDatasetAction`, antes de que el dataset ingrese al modo inserción. `gDetalle_OnAfterDatasetAction` solamente asigna los valores iniciales; no recorre el dataset durante la inserción, evitando el error `Dataset not in edit or insert mode`.

Para no agregar una consulta remota por cada línea insertada, el resultado de la validación de importaciones se conserva durante la edición del documento. La caché se reinicia al crear o mostrar otro documento, y la base de datos se consulta nuevamente antes de grabar.

La validación realizada en código comprende firmas, codificación, finales de línea, estructura de rutinas VB6 y revisión estática de los procedimientos. La ejecución contra una base de pruebas y la compilación completa de VB6 quedan como pasos obligatorios antes del despliegue productivo.

### 11.9 Resultado de validación técnica del 19/10/2026

- Se verificó que todas las rutinas modificadas tengan su correspondiente `End Sub`, `End Function` o `End Property`.
- Los formularios y la clase permanecen con finales de línea CRLF y codificación compatible con Windows-1252.
- Los procedimientos permanecen en UTF-8 sin BOM y LF; todos usan `CREATE OR ALTER`, `SET NOCOUNT ON`, columnas explícitas y `TRY/CATCH`.
- Los cinco scripts fueron analizados sin conexión mediante `TSql160Parser` de SQL Server Management Studio 22; el resultado fue cero errores sintácticos.
- No se modificaron archivos `.frx` ni la estructura de tablas.
- Se intentó compilar una copia aislada del proyecto con `VB6.EXE /make`, enviando la salida a una carpeta temporal para no reemplazar `Ventas.exe`. El compilador terminó antes del análisis con `Error al intentar tener acceso al registro del sistema`; por tanto, esta ejecución no confirma ni rechaza la compilación del código.
- Se confirmó que SQL Server Express está iniciado. La conexión integrada desde el proceso de validación no pudo crear el contexto SSPI y `sqlcmd` informó además una incompatibilidad de cifrado del cliente. No se utilizaron credenciales almacenadas ni se modificó ninguna base.
- LocalDB tampoco pudo iniciarse porque el mismo contexto no tiene acceso a su configuración del registro.

Antes de producción se debe compilar VB6 desde una sesión con los componentes registrados y desplegar los procedimientos en una base de pruebas con credenciales proporcionadas expresamente para ese ambiente.
