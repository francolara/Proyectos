# Análisis de parámetros de `frmDocVentas`

**Fecha:** 13/10/2026  
**Alcance:** `Formularios/frmDocVentas.frm`, `Modulos/mdlFunciones.bas` y `Modulos/mdlVariables.bas`.  
**Estado:** levantamiento técnico; no se modificó código ni base de datos.

## Conclusión

Sí conviene atacar este punto primero. El formulario consulta parámetros en tiempo de uso, no únicamente al iniciar. Se identificaron:

- 19 variables globales `STR_*` declaradas para la caché de parámetros.
- 17 de ellas usadas por `frmDocVentas`.
- 33 nombres de parámetros leídos mediante `traerCampo("Parametros", ...)` dentro del formulario.
- 15 nombres consultados mediante `leeParametro`, con 25 lecturas activas.

La sesión ya ejecuta `cargarParametrosSistema`, que usa una sola llamada a `spu_Parametros` por empresa. El problema es que la cobertura de esta caché es parcial y el formulario sigue emitiendo consultas unitarias en eventos y procesos de alto uso.

## 1. Carga actual de la caché

`cargarParametrosSistema` está en `Modulos/mdlFunciones.bas` (línea 1892). Se invoca desde `frmLogin` y también desde `frmPrincipal`. Obtiene todos los parámetros de `glsEmpresa` mediante:

```vb
EXEC spu_Parametros '<empresa>'
```

Después asigna variables públicas en `mdlVariables.bas`. Por lo tanto, no se debe volver a cargar toda la tabla al abrir cada documento: basta recargar al iniciar sesión, cambiar de empresa o guardar un mantenimiento de parámetros.

### Parámetros ya cacheados y usados por `frmDocVentas`

| Parámetro | Variable | Uso funcional principal | Lecturas aprox. |
|---|---|---|---:|
| `VISUALIZA_OC_FACTURA_PEDIDO` | `STR_VISUALIZA_OC_FACTURA_PEDIDO` | Muestra orden de compra en factura/pedido. | 1 |
| `IMPORTAR_DOCUMENTOS_ENTRE_EMPRESAS` | `STR_IMPORTAR_DOCUMENTOS_ENTRE_EMPRESAS` | Habilita importación entre empresas. | 1 |
| `EMPRESA_SUCURSAL_DOCUMENTOS_ENTRE_EMPRESAS` | `STR_EMPRESA_SUCURSAL_DOCUMENTOS_ENTRE_EMPRESAS` | Origen empresa/sucursal para importación. | 1 |
| `DESCRIPCION_AREA_O_UPP` | `STR_DESCRIPCION_AREA_O_UPP` | Etiqueta Área/UUPP. | 1 |
| `LONGITUD_IMPRESION_DETALLE_PRODUCTO` | `STR_LONGITUD_IMPRESION_DETALLE_PRODUCTO` | Longitud de detalle impreso. | 1 |
| `VENTA_ELECTRONICA` | `STR_VENTA_ELECTRONICA` | Visibilidad y flujo de comprobantes electrónicos. | 2 |
| `STOCK_POR_LOTE` | `STR_STOCK_POR_LOTE` | Manejo de lotes en detalle y stock. | 7 |
| `LIQUIDACIONES` | `STR_LIQUIDACIONES` | Lógica de liquidaciones. | 1 |
| `VALIDA_SEPARACION` | `STR_VALIDA_SEPARACION` | Validación de separación de mercadería. | 2 |
| `IGV`, `IGV_ANT`, `PERIODO_CAMBIO_IGV` | `STR_IGV`, `STR_IGV_ANT`, `STR_PERIODO_CAMBIO_IGV` | Determinación del IGV vigente. | 7 c/u |
| `CLIENTE_ANULA` | `STR_CLIENTE_ANULA` | Cliente usado en anulación. | 3 |
| `IMPORTA_ATENCIONES` | `STR_IMPORTA_ATENCIONES` | Habilita importación de atenciones. | 3 |
| `APRUEBA_PEDIDO_AUTOMATICO` | `STR_APRUEBA_PEDIDO_AUTOMATICO` | Aprobación de pedidos. | 1 |
| `SUCURSAL_PARTIDA_UBIGEO`, `SUCURSAL_PARTIDA_DIRECCION` | variables homónimas | Datos de partida de guía. | 2 / 1 |

## 2. Parámetros aún leídos directamente

Cada lectura listada abajo puede ser una llamada remota. Las ocurrencias están referidas a las líneas actuales de `frmDocVentas.frm`.

| Parámetro | Ocurrencias | Proceso | Acción propuesta |
|---|---:|---|---|
| `BLOQUEA_CHK_AFECTO` | 1 | Inicialización de la grilla. | Migrar de inmediato a `STR_BLOQUEA_CHK_AFECTO`; la variable ya existe y se carga. |
| `APRUEBA_PEDIDO_AUTOMATICO` | 3 | Grabación e importación de pedidos. | Migrar a `STR_APRUEBA_PEDIDO_AUTOMATICO`; ya existe y se carga. |
| `LIQUIDACIONES` | 3 | Listado y tratamiento de liquidaciones. | Migrar a `STR_LIQUIDACIONES`; ya existe y se carga. |
| `VALIDA_SEPARACION` | 3 | Validación del detalle. | Migrar a `STR_VALIDA_SEPARACION`; ya existe y se carga. |
| `CONTABILIDAD_ONLINE` | 3 | Generación y validación de asiento. | Añadir a caché de sesión. |
| `BLOQUEA_DETALLE` | 4 | Restricción de edición de grilla. | Añadir a caché de sesión; es un punto de alto impacto por interacción. |
| `SOLOVALIDADCTOMAX` | 2 | Validación de máximo de documentos. | Añadir a caché de sesión. |
| `AGRUPAPRODUCTOS` | 2 | Agrupación de detalle. | Añadir a caché de sesión. |
| `EVALUA_ITEM_DETALLE_PRODUCTO` | 2 | Evaluación del ítem de producto. | Añadir a caché de sesión. |
| `MONEDA_LINEA_CREDITO` | 2 | Control de línea de crédito. | Añadir a caché de sesión. |
| `CONCEPTO_VENTA_GRATUITA` | 1 | Venta gratuita. | Añadir a caché de sesión. |
| `VIZUALIZA_CODIGO_RAPIDO` | 1 | Identificación de producto. | Añadir a caché de sesión. |
| `MODIFICA_REGDOC_DESPROD` | 1 | Edición de registro/documento. | Añadir a caché de sesión. |
| `VALIDA_REFERENCIA` | 1 | Referencias documentarias. | Añadir a caché de sesión. |
| `PRODUCTO_ANULA` | 1 | Anulación. | Añadir a caché de sesión. |
| `IMP_FORMATO_PEDIDO` | 1 | Impresión de pedido. | Añadir a caché de sesión. |
| `CONCEPTO_AUTOMATICO_SALIDA`, `ALMACEN_AUTOMATICO_DESTINO`, `CONCEPTO_AUTOMATICO_INGRESO` | 1 c/u | Generación automática de movimientos. | Añadir a caché de sesión. |
| `CODIGO_PRODUCTO_APIMAS` | 1 | Facturación entre empresas/APIMAS. | Añadir a caché de sesión. |
| `CONTABILIDAD_ONLINE_VB`, `ORIGEN_CONTABLE` | 1 c/u | Integración contable. | Añadir a caché de sesión. |
| `PORCENTAJE_FACTURACION` | 1 | Cálculo de importación/facturación. | Añadir a caché de sesión. |
| `AADE_DESCRIPCION_AL_PRODUCTO` | 1 | Actualmente forzado a `"N"`. | No cachear hasta reactivar la funcionalidad. |
| `GENERA_VALE_FORMULA` | 1 | Actualmente inactivo (`If "N" = "S"`). | No cachear hasta reactivar la funcionalidad. |
| `REDONDEA_DETRACCION_COMPRAS` | 1 | Consultado por `CIdEmpresaCli`, no necesariamente la empresa de sesión. | No globalizar: caché por empresa consultada o devolverlo dentro del paquete de datos de la empresa cliente. |

## 3. Parámetros solicitados por `leeParametro`

`leeParametro` también consulta la tabla de parámetros de forma individual. Esta es la fuente más clara de viajes repetidos dentro de cálculos y eventos de grilla.

| Parámetro | Lecturas | Uso | Acción propuesta |
|---|---:|---|---|
| `EXTRAE_PRECIO_COMPRAS` | 4 | Fuente de precios de productos. | Caché de sesión. |
| `EXTRAE_PRECIO_ULTIMA_VENTA` | 4 | Fuente de precios de productos. | Caché de sesión. |
| `VENTA_ELECTRONICA` | 2 | Electrónica y validación de serie. | Usar solo `STR_VENTA_ELECTRONICA`. |
| `FACTURACION_ENTRE_EMPRESAS` | 3 | Procesos entre empresas y contabilidad. | Caché de sesión. |
| `VALIDA_FACTURADO_HCOSTO` | 2 | Validación de hoja de costo. | Caché de sesión. |
| `PORCENTAJE_IVAP` | 4 | Cálculo de IVAP. | Caché de sesión, convertido una sola vez a número. |
| `SOLOGUIAMUEVESTOCK` | 1 | Movimiento de stock. | Usar `glsSoloGuiaMueveStock`, que ya está cargado como Boolean. |
| `DNSCONTABILIDAD` | 1 | Nombre/base de datos contable para el `JOIN`. | Caché de sesión; validar que sea un identificador permitido antes de armar SQL. |
| `TIPO_IGV` | 1 | Tipo de IGV. | Caché de sesión. |
| `AYUDA_CENTRO_COSTO_DETALLE` | 1 | Ayuda de centro de costo. | Caché de sesión. |
| `CARPETA_XML_VE` | 1 | Apertura de PDF/XML electrónico. | Caché de sesión. |
| `VALIDA_PERMISOS` | 1 | Validación de permisos de nota de crédito. | Caché de sesión. |
| `FORMAPAGO_PEDIDO` | 1 | Valor por defecto de pedido. | Caché de sesión. |
| `CODIGO_SERVICIO_ATENCIONES` | 1 | Registro de atenciones. | Caché de sesión. |

## 4. Diseño recomendado

No recomiendo seguir agregando variables públicas sueltas sin estructura. En VB6, la alternativa gradual y segura es ampliar la carga única actual y mantener un acceso centralizado:

1. `cargarParametrosSistema` sigue siendo el único punto que consulta todos los parámetros de `glsEmpresa`.
2. Almacenar los valores en una caché en memoria indexada por nombre, además de conservar temporalmente las variables existentes para no romper otros formularios.
3. Exponer `ObtenerParametroSesion(nombre)` que no consulte la base: solo lee la caché y devuelve un error controlado si falta el nombre.
4. Crear conversiones centralizadas para bandera (`S/N`, `1/0`), decimal y texto. Evita que cada formulario interprete de forma distinta un valor nulo o vacío.
5. Recargar la caché únicamente cuando cambie empresa/sucursal o después de guardar el mantenimiento de parámetros.
6. Para `REDONDEA_DETRACCION_COMPRAS`, usar una segunda caché indexada por empresa, pues la consulta apunta a `CIdEmpresaCli`.

## 5. Orden de implementación en `frmDocVentas`

1. Reemplazar los cuatro parámetros que ya cuentan con variable global: `BLOQUEA_CHK_AFECTO`, `APRUEBA_PEDIDO_AUTOMATICO`, `LIQUIDACIONES` y `VALIDA_SEPARACION`.
2. Reemplazar los valores ya disponibles con distinta nomenclatura: `VENTA_ELECTRONICA` y `SOLOGUIAMUEVESTOCK`.
3. Añadir a la carga de sesión el paquete usado por la grilla y precios: `BLOQUEA_DETALLE`, `EXTRAE_PRECIO_COMPRAS`, `EXTRAE_PRECIO_ULTIMA_VENTA`, `PORCENTAJE_IVAP`, `AGRUPAPRODUCTOS` y `VIZUALIZA_CODIGO_RAPIDO`.
4. Añadir el paquete de grabación y contabilidad: crédito, referencias, aprobación, electrónica, stock y contabilidad.
5. Medir antes/después: número de llamadas SQL y tiempo de abrir, editar una línea, recalcular y grabar un documento por tipo.

## 6. Riesgos funcionales a controlar

- No convertir a global por sesión un valor que depende de una empresa distinta a `glsEmpresa`.
- Mantener exactamente las convenciones actuales (`"S"`, `"N"`, `"1"`, `"0"`) durante la primera migración; normalizarlas recién con pruebas.
- Invalidar la caché tras cambios de parámetros, no en cada `Form_Load`.
- Validar Factura, Pedido, Guía, Nota de Crédito, Cotización y Separación, porque habilitan rutas distintas.
- No tocar `frmDocVentas_Mostrar_V` ni `frmDocVentas_Vista` todavía; primero estabilizar y medir el formulario operativo, luego replicar mediante una capa común.
