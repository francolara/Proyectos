# Análisis: producto y precio en `frmDocVentas`

**Fecha:** 13/10/2026  
**Alcance:** selección y edición de una línea del detalle de venta.  
**Estado:** análisis previo a implementación; no modifica las reglas de cálculo.

## Resultado

Al seleccionar un producto, `gDetalle_OnEdited` ejecuta varias consultas independientes para completar una sola fila. En el caso normal de precio de lista son **seis accesos**; pueden ser siete o más si el producto es combo, tiene IVAP o se valida descuento.

La mejora propuesta es devolver el paquete completo de producto/precio en un único procedimiento almacenado. Los cálculos locales de moneda, IGV, IVAP, descuentos y totales se conservan en VB6 en la primera etapa para no cambiar resultados funcionales.

`DatosPrecioCompras` y `DatosPrecioUltimaVenta` están deshabilitados mediante comentarios en el formulario; no forman parte de este alcance ni del procedimiento propuesto.

## Implementación inicial

Se incorporó `dbo.spu_Docventas_ObtieneProductoPrecio`, utilizado al editar `idProducto` en `gDetalle_OnEdited`. En una llamada obtiene el producto activo, su marca, afectación IGV, moneda de la lista, unidad, factor, precio de lista y afectación de descuento especial.

Por seguridad funcional, `spu_Trae_Costo_Precio_Cotizacion` se conserva como una llamada separada: sigue aplicando la conversión y posible reemplazo por el precio especial del cliente exactamente como antes.

## Flujo actual al seleccionar producto

| Orden | Origen actual | Datos obtenidos |
|---:|---|---|
| 1 | Consulta directa a `Productos` | Producto canónico, descripción, unidad de venta y código rápido. |
| 2 | `DatosProducto` | Fabricante, marca, descripción de marca, afectación IGV y tipo de producto. |
| 3 | `Listaprecios` | Moneda de la lista. |
| 4 | `DatosPrecio` | Unidad, factor y valor de venta según la lista de precios. |
| 5 | `spu_Trae_Costo_Precio_Cotizacion` | Proveedor, costos, último precio por cliente y precio de pedido relacionado. |
| 6 | `Productos.afectoDctoEspecial` | Aplica descuento especial del cliente. |
| 7 | `LlenaFormulas` si es combo | Componentes y total calculado del combo. |

Después se ejecutan `procesaMoneda` y `calculaTotalesFila`. Actualmente ambas pueden volver a consultar `Productos.IndAfectoIvap` cuando corresponde IVAP.

## Rutinas y responsabilidad

| Rutina | Línea aprox. | Acceso principal |
|---|---:|---|
| `gDetalle_OnEdited` | 9118 | Orquesta la selección/edición de la línea. |
| `DatosProducto` | 14187 | Consulta producto y marca. |
| `DatosPrecio` | 14149 | Obtiene precio de lista y presentación. |
| `spu_Trae_Costo_Precio_Cotizacion` | Base de datos | Recupera costos y referencias de cotización. |
| `procesaMoneda` | 14216 | Convierte moneda y calcula IGV/IVAP unitario. |
| `calculaTotalesFila` | 14356 | Calcula descuentos, totales e IVAP neto. |

## Procedimiento unificado propuesto

Nombre sugerido: `dbo.spu_DocVentas_ObtenerProductoPrecio`.

### Parámetros de entrada

| Parámetro | Uso |
|---|---|
| `@IdEmpresa` | Empresa de la sesión. |
| `@IdSucursal` | Sucursal de la sesión. |
| `@IdDocumento` | Tipo de documento de venta. |
| `@IdLista` | Lista de precio. |
| `@IdProductoBusqueda` | Código ingresado: producto, fabricante o código rápido. |
| `@IdCliente` | Cliente para precio/cotización. |
| `@IdMonedaDocumento`, `@TipoCambio` | Contexto de costo/cotización. |

### Resultado esperado: una fila

- Identidad: `IdProducto`, `GlsProducto`, `CodigoRapido`, `IdUMVenta`.
- Producto: `IdFabricante`, `IdMarca`, `GlsMarca`, `AfectoIGV`, `IndAfectoIvap`, `AfectoDctoEspecial`, `IdTipoProducto`.
- Precio: `IdMonedaLista`, `IdUM`, `GlsUM`, `Factor`, `VVUnit`.
- Cotización: `Proveedor`, `CostoS`, `CostoD`, `CostoSInc`, `VvUnitario`, `Moneda`, `VvUnitarioG`.
- Combo: indicador de si existen componentes. Los componentes se mantendrán en una segunda consulta únicamente si aplica.

## Reglas que deben mantenerse

1. La búsqueda acepta `IdProducto`, `IdFabricante` o `CodigoRapido`.
2. El precio base se obtiene exclusivamente de la lista de precios mediante `DatosPrecio`.
3. Para servicios (`IdTipoProducto = '06002'`) se mantiene la regla actual de precio.
4. Un precio de cotización del cliente puede reemplazar el precio base cuando es mayor que cero.
5. El cálculo de moneda, IGV, IVAP y descuentos sigue en VB6 inicialmente.

## Intervención por etapas

1. Crear el SP unificado y probarlo contra las seis consultas actuales para una muestra de productos, servicios y combos.
2. Crear una función VB6 que consuma el resultado y complete los valores de salida actuales.
3. Reemplazar solo la rama `idProducto` de `gDetalle_OnEdited`.
4. Eliminar las consultas repetidas de `IndAfectoIvap` pasando el dato ya recuperado a `procesaMoneda` y `calculaTotalesFila`.
5. Medir número de llamadas y tiempo de alta de una línea antes/después.

## Riesgos a probar

- Productos sin precio, sin presentación o inactivos.
- Servicios y productos con IVAP.
- Producto con descuento especial.
- Cotización, pedido, factura, guía y nota de crédito.
- Combos y fórmulas de producto.
