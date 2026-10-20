# Análisis: detalle, totales y validaciones de `frmDocVentas`

**Fecha:** 13/10/2026  
**Alcance:** `gDetalle_OnEdited`, `procesaMoneda`, `calculaTotalesFila` y `calcularTotales`.

## Resultado del análisis

Los cálculos de precio, IGV, IVAP, descuentos y totales dependen del contenido temporal de `gDetalle` y de controles de la pantalla. Deben mantenerse en VB6 para no cambiar su secuencia funcional.

El principal problema de latencia era la lectura repetida de `Productos.IndAfectoIVAP`: se ejecutaba al cambiar precio, al recalcular una fila y una vez por cada fila cada vez que se recalculaban los totales. Ahora se mantiene una caché por instancia de `frmDocVentas`; el valor se precarga desde `spu_Docventas_ObtieneProductoPrecio` y solo consulta la base si se trata de una línea existente que aún no fue cargada en memoria.

Las validaciones de descuento también reutilizan una caché de `MaxDcto`, indexada por lista y producto. El límite se recibe desde el mismo SP al cargar un producto; si se edita una línea preexistente, se consulta una única vez y se conserva durante el documento.

## Flujo por edición de columna

| Evento de detalle | Reglas locales | Acceso a base que se mantiene |
|---|---|---|
| `idProducto` | Asigna producto, unidad, precio, IGV, descuento especial y recalcula la fila. | SP de producto/precio y SP de cotización. |
| `VVUnit` / `PVUnit` | Convierte moneda, calcula IGV/IVAP, valida hoja de costo y recalcula. | Centro de costo y facturación por hoja de costo, solo cuando aplica. |
| `Cantidad` | Recalcula importes, talla/peso y totales. | Consulta de talla/peso del producto. |
| `PorDcto`, `DctoVV`, `DctoPV` | Valida autorización y límite de descuento; recalcula la fila. | Unidad de venta y `MaxDcto` de la lista. |
| Centro de costo | Completa datos de presentación y controla el saldo facturable. | Centro de costo y procedimiento de facturación. |
| Costos y margen | Recalcula precio y totales con el margen indicado. | No requiere una consulta adicional de precio. |

## Límites de la etapa actual

1. `calcularTotales` permanece local porque recorre filas no guardadas y aplica descuento global de la pantalla.
2. Las validaciones de descuento y hoja de costo permanecen en línea: pueden bloquear o solicitar autorización.
3. La siguiente mejora candidata es agrupar la consulta de centro de costo y el saldo facturable en un procedimiento específico, sin cambiar sus bloqueos ni mensajes.
