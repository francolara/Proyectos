# Etapas 3 y 4 — implementación y entrega local

03/11/2026. Implementadas y verificadas localmente sobre MVC, C#, ADO.NET y procedimientos almacenados. Se entregan scripts para ejecución manual; no se abrió una conexión SQL Server, no se creó la base y no se desplegó la aplicación.

## Cobertura de aceptación

| Requisito | Implementación | Evidencia local |
|---|---|---|
| Clientes, tiendas, contactos y datos HYT por cliente | MaintenanceDefinitions, MasterController y cuatro familias de SP | Listados y fichas responden por HTTP; validación de nombres/documentos y vínculo de tienda-cliente en Guardar |
| Portal documental por cliente | UrlCargaDocumentos, validación HTTP(S), enlace externo seguro en ficha | Validación rechaza enlaces ejecutables; enlace exige esquema y ausencia de credenciales |
| Unidades, zonas, especialidades y tipos de servicio | Cuatro familias de mantenimiento con baja lógica | Rutas y formularios HTTP; FK preservadas y sin DELETE del maestro |
| Materiales con códigos alfanuméricos y varias zonas | ServicioMaterial y MaterialZona | Prueba de código con ceros; Chrome conserva selecciones Sur/Centro entre páginas |
| Búsqueda y ayudas de veinte | SP Listar con OFFSET/FETCH 20 y total; popup JavaScript compartido | HTTP verifica parámetros; Chrome cambia páginas y selecciona referencia con retorno de foco |
| Listado/ficha y retorno a filtros | _Grid y vistas Master | Pruebas de búsqueda/página/returnUrl y rechazo de retorno externo |
| Auditoría y fechas responsables | FechaRegistro/Actualizacion, EventoAuditoria y Get con historial paginado | HTTP de trazabilidad, codificación HTML y horario de Perú; SP Guardar audita en transacción |
| Edición concurrente | VersionFila original, bloqueos en SQL, mensajes de conflicto | HTTP conserva el texto editado al rechazar guardado concurrente; revisión offline de las condiciones SQL |
| Tarifarios y versiones | Tarifario/Version/Partida, solo BORRADOR editable | HTTP rechaza POST directo a versión publicada |
| PEN inicial, USD opcional y tipo IGV por registro | Catálogos internos, valores predeterminados y precios decimales | Validación mantiene USD NULL; Chrome aplica moneda/IGV/zonas y conserva USD vacío |
| Predeterminados no reescriben partidas | Copia a nueva partida desde versión; ayuda explícita para códigos del material | Chrome comprueba códigos con ceros; Guardar versión no actualiza partidas |
| Publicación y nueva versión | Publicar y Clonar con permisos, versión original y transacción | HTTP de confirmación, POST y destino nuevo; scripts copian zonas/partidas y auditan origen/destino |
| Vigencias IGV en tabla | TaxRateController, SqlTaxRateRepository y dos SP | HTTP de tasa válida/inválida y conflicto de vigencias; alta y cierre anterior auditados |
| Identidad visual y adaptación | Estilos HYT y calendario común | Chrome escritorio y móvil; capturas de materiales revisadas visualmente, sin desbordamiento |
| Errores y permisos | Políticas, antifalsificación y MaintenanceUnavailableFilter | Acceso sin permiso rechazado; error 503 sin detalles internos; formulario inválido no guarda |
| Scripts y paquetes | Manifiestos etapas 3/4 y SHA256 | 109 fuentes, 46 tablas, 103 relaciones; sintaxis, columnas resolubles y huellas verificadas offline |

## Entregables

- `entregables/ETAPA_3_HYT.sql`: 30 fuentes incrementales.
- `entregables/ETAPA_4_HYT.sql`: 14 fuentes incrementales.
- `entregables/SHA256_ETAPA_3.json` y `SHA256_ETAPA_4.json`.
- `GUIA_ETAPAS_3_Y_4.md`: instalación manual y comprobaciones en SQL.
- `entregables/preview/material-escritorio.png` y `material-movil.png`: capturas de datos de prueba; no contienen datos reales.

## Verificación y límite de evidencia

Compilación sin errores ni advertencias. 69 comprobaciones C#/HTTP aprobadas sobre el host MVC con repositorios de prueba; Chrome sobre el mismo host MVC usando repositorios aislados, sin SQL. Cinco regresiones demuestran que el validador offline detecta sintaxis inválida, FK inexistente, pérdida de unicidad de guía, SELECT * y columnas inexistentes.

La validación offline no demuestra ejecución, permisos efectivos ni concurrencia real de SQL Server. Estas comprobaciones quedan expresamente reservadas al usuario, conforme a su instrucción de entregar scripts y no ejecutar la base. Tampoco se importaron los Excel: se conserva el tarifario como referencia; cualquier carga de antecedentes requiere revisar la correspondencia de columnas y duplicados.

La siguiente etapa funcional es presupuestos (etapa 5), con conversiones IGV y formatos HTML anterior/nuevo. Dashboard y seguimiento documental mantienen las etapas acordadas.
