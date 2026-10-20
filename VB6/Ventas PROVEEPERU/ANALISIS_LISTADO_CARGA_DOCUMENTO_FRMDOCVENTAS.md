# Análisis: listado y carga de documento en `frmDocVentas`

**Fecha:** 13/10/2026  
**Alcance:** `listaDocVentas`, `ListaDetalle`, `gLista_OnChangeNode`, `gLista_OnDblClick` y `mostrarDocVentas`.

## Flujo actual

1. `listaDocVentas` invoca `spu_Docventas_Listacab` y enlaza la grilla de cabeceras.
2. Al cambiar la fila, `gLista_OnChangeNode` ejecuta `ListaDetalle` para mostrar la previsualización mediante `spu_Docventas_ListaDet`.
3. Al abrir el documento, `mostrarDocVentas` obtiene por separado:
   - Cabecera: `spu_Docventas_Mostrar_Docventas_cab`.
   - Detalle completo: `spu_Docventas_Mostrar_Docventas_Det`.
   - Referencias: `spu_Docventas_Mostrar_Docventas_Ref`.
   - Lotes y liquidaciones, solo si los parámetros correspondientes están activos.

## Mejora aplicada

`ListaDetalle` conserva la clave documento/serie/número que ya está visible. Si el evento se dispara nuevamente para la misma fila después de enlazar el listado, no vuelve a ejecutar `spu_Docventas_ListaDet`. La clave se reinicia al renovar el listado.

## Carga consolidada aplicada

`dbo.spu_Docventas_CargarDocumento` entrega, en este orden, cabecera, detalle y referencias. `mostrarDocVentas` los consume mediante `NextRecordset` y conserva los recordsets temporales que ya alimentan el formulario y las grillas.

Esto reemplaza tres viajes de red por una sola llamada al abrir un documento. No se usa un `JOIN` único, por lo que la cabecera no se repite por ítem ni las referencias multiplican el detalle.

Lotes y liquidaciones permanecen condicionales en esta primera consolidación, ya que solo aplican a configuraciones específicas.
