# Mapeo del formato nuevo HYT

Referencia: C:/Users/Franco Lara/Desktop/hytpresupuesto.pdf, una página A4. Revisión: 06/11/2026.

| Campo del ejemplo | Origen en HYT | Situación |
|---|---|---|
| Logo, empresa y RUC | Logo HYT y EmpresaSnapshotJson | Mapeado |
| Línea de servicios, R.P.C y correo | SubtituloCotizacion de Configuración HYT | Mapeado: texto completo libre, incluidos los dos correos, centrado en impresión |
| CLIENTE / DIRECCION | ClienteNombre / ClienteDireccion de revisión | Mapeado; dirección conserva ubigeo |
| PPTO N° / FECHA | NumeroPresupuesto visible / FechaEmision | Mapeado |
| TRIMESTRE | Trimestre | Mapeado |
| Zona / precios válidos | ZonaTexto / VigenciaTexto | Mapeado; vigencia es texto libre, no un período calculado |
| NOMBRE proveedor | RazonSocial de EmpresaSnapshotJson | Mapeado |
| CODIGO SAP proveedor | CodigoSAPProveedor de revisión, obtenido del cliente | Mapeado |
| CONTRATO MARCO / ESPECIALIDAD | ContratoMarco / Especialidad | Mapeado |
| N° COTIZACION / FECHA DE EJECUCION | NumeroCotizacionCliente / FechaEjecucion | Mapeado |
| TIENDA | TiendaCodigo de revisión, nombre como respaldo | Mapeado |
| CENTRO SAP FICO | CentroSAPFICO por ítem, obtenido siempre de la tienda al guardar | Mapeado; requiere dato configurado en la tienda |
| N° NECESIDAD / TIPO DE SERVICIO | NumeroNecesidad / TipoServicio del ítem | Mapeado |
| POSICION | Posicion del ítem | Mapeado: código interno del material (CodigoInterno); se obtiene desde maestro al guardar y se imprime también en presupuestos previamente registrados |
| CODIGO MATERIAL SAP / DESCRIPCION | CodigoMaterialSAP / Descripcion del ítem | Mapeado |
| UNIDAD / CANT / PRECIO / SUBTOTAL | Unidad / Cantidad / PrecioUnitario / BaseImponible | Mapeado; precios netos y dos decimales |
| Fila roja de agrupación | Grupo del ítem, si está registrado | Mapeado; no imprime DescripcionTrabajo y sustituye el error #N/A del Excel con el grupo real o espacio vacío |
| SUBTOTAL / IGV / TOTAL | Subtotal / Impuesto / Total de revisión | Mapeado; tasa histórica y moneda de presupuesto |
| Consideraciones | Consideraciones de revisión | Mapeado; no agrega notas contradictorias ni repite despedidas que ya estén escritas |
| Firmante y cargo | Firmante/CargoFirmante de revisión o predeterminado HYT | Mapeado |
| Firma escaneada | FirmaPresupuestoJson en Configuración HYT → imagen privada R2 | Implementado; requiere que el usuario cargue su JPG/PNG |

Los campos internos continúan guardados y se omiten en el formato nuevo. Los errores #N/A y los ceros de filas de ejemplo no son datos que deban copiarse. Se conserva una rejilla mínima de doce filas, en blanco donde no hay ítems. Las descripciones largas y varios ítems pueden aumentar las páginas, con cabeceras repetidas y sin truncar el detalle.

Actualización 06/11/2026: los tres puntos quedan resueltos. Posición usa el código interno del material; Centro SAP FICO proviene de tienda y el subtítulo configurable admite el texto completo con ambos correos. El archivo PDF no se utiliza para extraer ni cargar automáticamente una firma.

Ajuste de presentación 06/11/2026: Fecha reduce su ancho 40 % y el espacio se asigna al número de presupuesto. Cuadros de trimestre/vigencia/proveedor con valores centrados, ancho de 39 mm (antes 30 mm) y altura común de 5 mm. Descripción del trabajo se conserva interna y se omite en ambos formatos.
