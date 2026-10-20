# Análisis del botón Grabar — FrmDocVentas

Fecha: 13/10/2026. Actualizado: 15/10/2026.

## Alcance

Se revisó `Grabar` en `Formularios/frmDocVentas.frm` y la persistencia que delega a `Clases/clsDocVentas.cls`.

## Secuencia actual

1. Valida cierre de mes, campos obligatorios, detalle y referencias.
2. Valida centro de costo, stock por lote y stock/saldos cuando corresponda.
3. Determina si se inserta o modifica, genera correlativo si falta y valida la referencia guía/factura.
4. Abre una transacción en la conexión existente.
5. Graba cabecera, detalle, referencias, lotes, liquidaciones y vales mediante `EjecutaSQLFormDocVentas` o `EjecutaSQLFormDocVentasEmpresas`.
6. Ejecuta operaciones complementarias: factura asociada, cuentas por cobrar, línea de crédito, atribución de nota, pedidos y aprobaciones.
7. Confirma o revierte la transacción y actualiza el listado.

## Optimización aplicada

Durante una grabación se almacenan por formulario los datos maestros que pueden repetirse:

- Tipo de producto: evita una consulta por cada comprobación de lote del mismo producto.
- Indicador de transferencia de documento/serie: se consulta una vez para alta o modificación.
- Unidad de negocio del centro de costo: evita repetir la misma lectura si el detalle usa la misma hoja de costo.

La caché solo vive mientras el formulario está abierto. No reemplaza validaciones transaccionales ni altera los datos grabados.

## Primera migración a procedimiento

Se creó `dbo.spu_Docventas_PostGraba` y `clsDocVentas` lo invoca al finalizar la persistencia principal. Reemplaza dos viajes a la base de datos por uno solo:

- Guarda `ObsDocVentas` y `ObsDocVentas2`.
- Calcula desde `DocVentasDet` y actualiza `TotalTallaPeso`.

El procedimiento se ejecuta dentro de la transacción que abre `FrmDocVentas`, por lo que cualquier error conserva el comportamiento de reversión existente.

## Escrituras posteriores al detalle

Se creó `dbo.spu_Docventas_GrabaPostDetalle`. Recibe las referencias activas en XML y, en una sola llamada, reemplaza las referencias del documento y actualiza el vendedor de campo del cliente. El recálculo de `CantidadImp`, `EstDocImportado` y `DocReferenciaDet` no fue modificado.

Los bloques de fórmulas, liquidaciones y detalle por lote fueron comentados en `EjecutaSQLFormDocVentas` porque ya no están en uso. Se conserva la generación normal de vales que no depende de lote ni fórmula.

## Generación del vale normal

Se creó `dbo.spu_Docventas_GeneraVale`. Reemplaza las consultas y escrituras de `generaValeDocVentas` que eliminaban o generaban `ValesCab`, `ValesDet`, `DocReferencia` y el estado del documento. El procedimiento genera el correlativo dentro de SQL Server con bloqueo de concurrencia y ejecuta esas operaciones como una unidad.

La reversa y aplicación de stock/costo siguen en `actualizaStock` de VB6, antes y después del procedimiento, respectivamente. Con ello se conserva el cálculo actual de existencias mientras la grabación del vale pasa a una sola llamada a la base. El vale de ingreso de transferencias queda pendiente de su procedimiento específico.

## Punto crítico pendiente

La mayor cantidad de viajes a la base está dentro de `clsDocVentas.EjecutaSQLFormDocVentas`: borra e inserta detalle/referencias y realiza operaciones complementarias por fila. Convertirlo directamente en un único procedimiento implicaría transportar el detalle, referencias, lotes y liquidaciones como colecciones estructuradas y conservar todas las reglas de stock, vales e importaciones.

Por seguridad, esa migración debe realizarse como una segunda etapa con procedimientos específicos por bloque y pruebas funcionales de venta, guía, nota de crédito, pedido, lote y liquidación. La transacción actual se conserva sin cambios.
