# Tarifario 2026-2028: análisis inicial

Fecha: 03/11/2026. Fuente: `C:/Users/Franco Lara/Desktop/NUEVO TARIFARIO  2026-2028.xlsx`, hoja `Tarifario INFRA-IISS`. Revisión de estructura y valores almacenados del XLSX, sin modificarlo ni importar sus registros.

Se identifican 1.162 filas con descripción después de los encabezados. Esto no equivale todavía a 1.162 materiales únicos validados.

## Columnas y tratamiento

| Columna | Contenido | Tratamiento propuesto |
|---|---|---|
| A | Posición | Texto de libre registro; no usar como clave única |
| B | Código material SAP | Texto alfanumérico; admite repeticiones según contexto |
| C | Descripción | Conservar detalle completo |
| D | Unidad | Vincular con mantenimiento de unidades de medida |
| E | Precio PEN para Sur-Centro-Norte | Asociar a la versión del tarifario y zonas confirmadas |
| G, H, I | Tipo de servicio, número de necesidad y especialidad | Solo tienen valores en las primeras nueve filas de datos; validar si son ejemplos/listas auxiliares o atributos de esas partidas; no extenderlos automáticamente al resto |

El código interno no aparece como columna independiente: se incorpora por requisito del usuario. Su asignación a antecedentes se acordará antes de la carga.

## Calidad y decisiones de carga

- Unidades encontradas: UND, M2, ML, PTO, M3, CAJA, KG, MES, DIA, BLS, GLN y SEM.; también `UND ` y `m2`, que requieren equivalencias con UND y M2 previa revisión.
- En las filas finales se repiten posición 1152 y código SAP 8023487 con descripciones e importes diferentes. No deduplicar solo por esos campos ni imponer una clave única que pierda partidas.
- El encabezado agrupa Sur, Centro y Norte sobre una sola columna de precio. Confirmar aplicación común y vincular explícitamente las zonas del mantenimiento.
- El periodo 2026-2028 no determina por sí solo fechas exactas de inicio y fin; confirmar antes de activar la versión.
- La regla de IGV sigue pendiente: no puede deducirse solo de este tarifario.

## Alcance confirmado del sistema

Mantenimientos de unidades y zonas; material con posición, código interno y código SAP alfanuméricos de libre registro; una unidad seleccionada y múltiples zonas por material. Componentes de listado, edición, consulta y auditoría conforme al plan de HYT. La carga de datos requerirá vista de revisión y conciliación de cantidades, códigos e importes antes de incorporarlos a la base.
