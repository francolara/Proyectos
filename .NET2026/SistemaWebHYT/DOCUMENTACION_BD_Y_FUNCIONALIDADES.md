# Base de datos y funcionalidades de Sistema Web HYT

03/11/2026. Esquema de etapas 1 a 4 instalado en SQL Server local por solicitud posterior del usuario. Las notas de entrega offline anteriores describen el estado previo a esta instalación.

## Ubicación y autoridad

Base propuesta: SistemaWebHYT. Fuentes SQL únicas en Basededatos/Dbsishyt/Tablas, StoreProcedure y Script, relativas a la raíz del repositorio, conforme a AGENTS.md. manifest.json registra el orden completo. No existe ya la ubicación inicial BaseDatos dentro del proyecto.

LEEME_DESPLIEGUE.md, junto al manifiesto, explica ejecución manual, cuentas, límites y configuración inicial. DICCIONARIO_BD_ETAPA_1.md de la raíz del proyecto enumera todas las columnas y restricciones. PLAN_DE_TRABAJO_Y_PROCESO.md conserva los requisitos funcionales; las últimas decisiones del usuario prevalecen sobre preguntas históricas.

## Estructura física preparada

| Área | Tablas |
|---|---|
| Catálogos internos | Moneda, TipoDocumento, TipoIGV, TasaIGV |
| Identidad | AspNetUsers, AspNetRoles, AspNetUserRoles, AspNetUserClaims, AspNetRoleClaims, AspNetUserLogins, AspNetUserTokens |
| Autorización y auditoría | Permiso, RolPermiso, EventoAuditoria |
| Configuración | ConfiguracionHYT, NumeracionInterna |
| Clientes | Cliente, Tienda, ContactoCliente, ProveedorCliente |
| Maestros | UnidadMedida, Zona, Especialidad, TipoServicio, ServicioMaterial, MaterialZona |
| Tarifarios | Tarifario, TarifarioVersion, TarifarioVersionZona, TarifarioPartida, TarifarioPartidaZona |
| Presupuestos | Presupuesto, PresupuestoRevision, PresupuestoPartida, ImpresionPresupuesto |
| Seguimiento | OrdenCompraCliente, Ejecucion, GuiaConformidad, GuiaPartida, SeguimientoCliente, ConfirmacionFacturacion, ReferenciaComprobanteExterno |
| Adjuntos | ArchivoAdjunto, OrdenCompraArchivo, GuiaArchivo, ConfirmacionArchivo |

## Reglas incorporadas en el esquema

- Una sola empresa: ConfiguracionHYT tiene clave restringida a 1; no hay partición EmpresaId en todas las entidades.
- Tienda específica obligatoria por presupuesto; FK compuesta garantiza tienda del mismo cliente.
- TipoAtencion restringido a INVITACION o EMERGENCIA. Una OC y una guía como máximo por presupuesto mediante claves únicas. Cada OC refiere una revisión del mismo presupuesto. La guía no exige OC y la hereda por su presupuesto cuando exista.
- NumeroRegistroInterno y NumeroOCCliente son diferentes; número externo, SAP, posición, número de guía y confirmación son texto. No se impone unicidad global a referencias externas o SAP.
- Catálogo Moneda por Codigo PEN/USD; documentos y tarifarios guardan esos códigos directamente. TipoDocumento se relaciona mediante 01/03. Solo los catálogos tienen clave única por código; las referencias son repetibles.
- Tarifario con versiones y predeterminados; zonas N:M por material, versión y partida. Zonas informativas en la ayuda de presupuesto. PrecioPEN y PrecioUSD decimales, USD admite NULL; sin conversión de divisas automática.
- TipoIGV separado de la tasa. Tasa 0.18 en TasaIGV, con intervalos FechaDesde inclusivo/FechaHasta exclusivo; nuevas vigencias cierran la anterior. Revisión conserva tasa aplicada con FK compuesta a vigencia/valor. Cálculos y redondeo se implementarán en etapa 4/5 sin revalorizar documentos históricos.
- Revisiones y partidas conservan datos de cliente/tienda, proveedor, contrato, precios, descripciones, unidades y referencias. Los cambios posteriores de maestros no deben sobrescribir esas copias.
- Cliente.UrlCargaDocumentos guarda el enlace del portal externo. No almacena contraseñas. Es independiente de los adjuntos privados de HYT.
- PDF/adjuntos con metadatos, ubicación privada, huella y vínculos concretos, sin binarios en consultas paginadas. Validación de archivo y visor se implementan en la etapa de interfaz.
- GuiaConformidad distingue registro de aceptación; ConfirmacionFacturacion conserva eventos y una confirmación vigente por OC. ReferenciaComprobanteExterno representa marcado manual Facturado, no emisión. Una referencia vigente por OC y FK a confirmación de esa OC. Trigger valida confirmación vigente, guía aceptada, trabajo terminado y OC no cancelada al marcar.
- Las fechas de documento son DATE y no se desplazan de día. Registro UTC DATETIME2(3), separado de fecha real del evento. Tablas editables tienen autoría y ROWVERSION; detalles/asociaciones se auditan como parte de la operación del documento padre.
- La impresión conserva revisión, formato ANTERIOR/NUEVO y versión del formato.

## Procedimientos y seguridad SQL preparados

| Procedimiento | Función |
|---|---|
| spu_RegistrarAuditoria | Anexa evento con usuario activo dentro de transacción de negocio; no se concede directamente al rol de aplicación |
| spu_ObtenerNumeroInterno | Reserva correlativo atómico dentro de la transacción del documento |
| spu_RegistrarTasaIGV | Configuración autorizada de nueva vigencia posterior, con cierre de la anterior y auditoría en la misma transacción |
| spu_ListarClientes | Búsqueda autorizada por código/nombre/documento, página SQL fija de 20 y total separado |

EventoAuditoria registra entidad, identificador, acción, estados, fecha del sistema, fecha del evento, motivo, cambios JSON y correlación. Trigger impide modificar/borrar el historial. El rol de conexión hyt_app_executor no puede escribir directamente; permisos EXECUTE se agregan individualmente. No se crean usuarios ni credenciales ni se concede acceso db_owner. La matriz inicial de roles es propuesta revisable en etapa 2.

La operación de usuario autenticado y permisos de archivo se implementará en el host; ningún endpoint debe aceptar UsuarioId arbitrario enviado por el formulario. Identity se integrará con stores ADO.NET y longitud de claves 128 coherente con las tablas.

## Requisitos de consultas posteriores

Todos los mantenimientos y ayudas: código/descripción y SQL de 20 filas. Documentos: búsqueda y Desde/Hasta por FechaRegistro, último mes móvil en America/Lima inicialmente, con límites UTC inclusivo/exclusivo y conservación de filtros al volver de ficha. Retornar total con los mismos filtros. Índices adicionales y procedimientos específicos se incorporarán con cada módulo; no se simulan consultas ya implementadas.

Dashboard con cifras reales y filtros: presupuestos, OC, sin confirmación, confirmadas por marcar y Facturado. Curvas interactivas, columnas y dona con categorías claras, sin duplicar revisiones/eventos. Totales por moneda, fechas de referencia y accesos a listados. Las consultas agregadas del dashboard se implementarán después de las operaciones, no forman parte de los cuatro SP de soporte.

## Verificación y límites

La etapa 1 entrega estructura y scripts. No implementa login, CRUD completo, importación de Excel, visualización de archivos, cálculos de presupuesto, dashboard conectado ni reportes: corresponden a etapas posteriores. SQL validado offline, no ejecutado. El usuario deberá revisar y ejecutar manualmente el paquete en su instancia de pruebas.

## Ampliación de etapa 2 — acceso

03/11/2026: sin cambios de tablas. El manifiesto acumulado ahora registra 65 scripts. manifest-etapa1.json conserva los 59 de la primera entrega y manifest-etapa2.json enumera seis incrementales: cinco SP de acceso y Script/010_PermisosAcceso.sql. El paquete de etapa 1 conserva sus fuentes/huellas; el de etapa 2 debe aplicarse después.

spu_ConsultarAcceso devuelve identidad, hash privado, estado, sello, roles y permisos solo a la cuenta de aplicación. spu_FinalizarAcceso valida de nuevo cuenta/sello/bloqueo, actualiza intentos o rehash y anexa inicio/rechazo en la misma transacción. Cinco fallos bloquean 15 minutos. spu_RegistrarSalida anexa cierre incluso para una cuenta desactivada.

spu_CrearAdministradorInicial es exclusivo de despliegue: bloqueo de aprovisionamiento, hash/sello generados por la herramienta, rol HYT-ADMIN y evento ADMINISTRADOR_INICIAL con login SQL. spu_RestablecerAcceso exige actor activo con USUARIOS.EDITAR, motivo, nuevo hash/sello y auditoría atómica. Ninguno se concede al rol web; no hay contraseña predeterminada. Las herramientas solo se ejecutarán por el usuario con una conexión administrativa autorizada.

AspNetUsers mantiene longitudes de 128 para Id y 256 para nombre normalizado, según esquema de etapa 1. La implementación usa PasswordHasher de Identity y repositorio ADO.NET/SP, no EF ni migraciones automáticas. La autoría de cada solicitud se toma de la identidad autenticada en el host. No se han ejecutado los scripts ni conexiones SQL.

## Tecnología definitiva: MVC

03/11/2026: por instrucción del usuario, Web usa ASP.NET Core MVC, controladores C#, ViewModels y vistas .cshtml; no usa Blazor ni Razor Pages. Infrastructure conserva ADO.NET y SP; no se modificaron tablas, scripts o procedimientos como consecuencia de esta migración de presentación. Las pruebas de acceso siguen usando un repositorio de prueba, sin ejecutar SQL Server.
## Mantenimientos de etapas 3 y 4 — 03/11/2026

El script incremental `011_EstadosMantenimientos.sql` añade `Activo BIT NOT NULL DEFAULT (1)` a `ProveedorCliente` y `TarifarioPartida`, conservando activos los registros existentes. Permite baja lógica y conserva sus referencias históricas. Aplicar después de las etapas 1 y 2; los paquetes anteriores permanecen intactos.

Los procedimientos `spu_Mant{Entidad}_Listar`, `_Obtener` y `_Guardar` implementan doce mantenimientos. Consultan páginas de veinte, trazabilidad e historial; guardan con comprobación de permisos, versión original y auditoría en la misma transacción. Las pantallas MVC están conectadas a estos procedimientos mediante ADO.NET.

`spu_MantCatalogos` devuelve monedas y tratamientos IGV a usuarios activos, sin permitir editarlos desde la aplicación. El repositorio `SqlMaintenanceRepository` consume exclusivamente procedimientos almacenados; su incorporación no abre conexiones durante la compilación.


## Versiones del tarifario — actualización 03/11/2026

`spu_MantTarifarioVersion_Publicar` exige permiso TARIFARIOS.EDITAR, versión original y estado BORRADOR; verifica tarifario activo, partidas activas, precio PEN presente y materiales/unidades activos. Cambia a VIGENTE con auditoría en la misma transacción. El precio USD continúa opcional. No modifica presupuestos ni otras versiones.

`spu_MantTarifarioVersion_Clonar` bloquea el tarifario padre y la versión original, asigna un número consecutivo y copia cabecera, zonas, partidas y zonas de partidas a un nuevo BORRADOR. Copia también valores nulos y registros inactivos para permitir su revisión explícita; la publicación vuelve a validar las partidas activas. Audita el origen, el destino y cada partida creada. No cambia el contenido de la versión fuente.

Los guardados de versión y las operaciones de publicación/copia usan el mismo orden de bloqueo (tarifario, versión). El guardado de partidas bloquea su versión y actualiza su rowversion, para detectar ediciones concurrentes antes de publicar/copiar. Las zonas históricas desactivadas pueden mantenerse; no se pueden agregar nuevas asociaciones a zonas desactivadas.

Los manifiestos incrementales etapa3/etapa4 y permisos 012/013 registran las nuevas fuentes. Los paquetes ETAPA_3_HYT.sql y ETAPA_4_HYT.sql incluyen huellas SHA256 y fueron analizados offline; no ejecutados en SQL Server. Aplicar en secuencia después de etapas 1 y 2. La prueba real contra motor SQL sigue reservada a ejecución manual del usuario.

## Vigencias IGV y cierre de etapas 3/4 — 03/11/2026

`spu_MantTasaIGV_Listar` devuelve hasta veinte vigencias con total, última fecha registrada y responsables/fechas obtenidos de auditoría. `spu_MantTasaIGV_Agregar` recibe fecha, valor decimal, motivo y última fecha esperada; serializa la tabla de vigencias, rechaza formularios desactualizados y fechas no posteriores, cierra la anterior e inserta la nueva, auditando ambos registros dentro de su transacción. Requiere CONFIGURACION.CONSULTAR/EDITAR según la operación. No cambia TasaIGV estructuralmente ni actualiza revisiones de presupuesto.

La entrega acumulada registra 109 fuentes. Etapa 3 contiene 30; etapa 4 contiene 14, incluidas vigencias IGV. Moneda y TipoIGV siguen como catálogos internos. GUIA_ETAPAS_3_Y_4.md documenta ejecución manual y verificaciones de motor pendientes para el usuario; ESTADO_ETAPAS_3_Y_4.md relaciona requisitos, implementación y evidencia local. No se ejecutó SQL Server.

## Instalación local autorizada — 03/11/2026

El usuario solicitó expresamente crear la base local con todas las tablas y SP, sustituyendo su restricción anterior de entregar solo scripts para esta instalación. Destino comprobado: localhost\SQLEXPRESS (DESKTOP-LLR14AJ\SQLEXPRESS), SQL Server Express 2022 versión 16.0.1000.6. La base SistemaWebHYT no existía antes de la operación.

Se aplicaron las 109 fuentes del manifiesto acumulado mediante autenticación integrada de Windows, UTF-8 y QUOTED_IDENTIFIER habilitado (`sqlcmd -E -C -I -f 65001 -b`). El primer intento se detuvo en el índice de AspNetUsers por la opción QUOTED_IDENTIFIER; se habilitó la opción y se reaplicaron las fuentes idempotentes, sin borrar objetos ni bases.

Verificación contra el motor: 46 tablas, 50 procedimientos, 103 claves foráneas; ninguna deshabilitada o sin confianza. Se refrescaron las definiciones de todos los SP mediante sp_refreshsqlmodule sin errores y DBCC CHECKCONSTRAINTS no informó violaciones. Catálogos comprobados: PEN/USD, 01/03, CON_IGV/SIN_IGV y tasa 0.180000 desde 03/11/2026. Existe el rol hyt_app_executor. AspNetUsers contiene cero usuarios: no se creó administrador ni contraseña de fábrica. No se modificó la configuración de conexión de la web, no se crearon logins SQL y no se desplegó a producción.

Esta comprobación valida instalación y resolución de objetos en el motor; no sustituye las pruebas funcionales completas de operaciones y concurrencia con usuarios autenticados.

## Etapas 5 y 6 en desarrollo — 03/11/2026

Nueva autorización del usuario: ejecutar directamente en SistemaWebHYT local los SP y cambios de estructura de estas etapas. Se confirmó localhost\SQLEXPRESS, con cero presupuestos, tarifas y usuarios antes de esta ampliación. No se autorizó ni ejecutó despliegue de producción.

Regla nueva confirmada: cantidades, precios unitarios e importes a dos decimales. BudgetAmounts implementa cálculo decimal por partida, redondeo comercial (mitades alejadas de cero) y suma de partidas. Convierte CON_IGV/SIN_IGV usando la tasa recibida, sin tasa constante y sin conversión cambiaria. La conversión de precio unitario se redondea a dos decimales antes de multiplicar; por eso distintos tratamientos pueden diferir un centavo cuando la conversión no es exacta. USD ausente deja importes incompletos/NULL, no cero. Cantidad/precio con más de dos decimales se rechazan.

Script/014_PresupuestosDocumentos.sql aplicado: PresupuestoRevision incorpora ZonaTexto, EmpresaSnapshotJson (con CHECK JSON), PoliticaCalculo y FechaAprobacion; PresupuestoPartida incorpora ZonasTexto, OrigenPrecio y MotivoPrecioManual; ImpresionPresupuesto incorpora ContenidoHTML y SHA256. Los campos HTML/snapshot se usarán al implementar la captura e impresión y todavía no contienen datos. CHECK nuevos exigen dos decimales en cantidades/precios de presupuestos, tarifas y guías, conservando los tipos de almacenamiento originales y su capacidad histórica. La tasa IGV conserva su precisión propia; no es un precio monetario.

Aplicados spu_Presupuesto_Listar, spu_Presupuesto_Obtener y spu_Presupuesto_AyudaTarifario, con permisos individuales en Script/015_PermisosEtapa5.sql. Listar recibe límites UTC y devuelve última revisión, estado, importes, OC, guía y fechas/responsables, veinte registros y total. Obtener devuelve snapshot/revisiones/partidas y auditoría de veinte. La ayuda ofrece tarifas publicadas de cliente o genéricas por fecha y moneda, mantiene precio USD nulo y zonas informativas sin filtro/prioridad por zona. Fuentes registradas en manifest-etapa5.json y manifiesto acumulado (114 fuentes).

Prueba real contra SQL Server: consultas ejecutadas dentro de transacción con actor efímero y rollback, sin usuarios de prueba persistentes. También se aprobaron pruebas de cálculo, precisión, USD ausente y bloqueo de mezcla de monedas, junto con las pruebas existentes de MVC. Todavía faltan guardado/transiciones/revisiones del presupuesto, formularios, ambos HTML, seguimiento/OC/guía/confirmación/adjuntos y pruebas integrales de etapas 5 y 6.

### Guardado y estados de presupuesto — avance 03/11/2026

Aplicados en la base local `spu_Presupuesto_Guardar` y `spu_Presupuesto_CambiarEstado`; la etapa 5 parcial registra siete fuentes y el acumulado 116. `PresupuestoRevision.DescripcionTrabajo NVARCHAR(1500) NULL` conserva la descripción por revisión para que una nueva revisión no cambie una impresión histórica.

Guardar administra una transacción propia, verifica PRESUPUESTOS.EDITAR, ambas versiones originales (cabecera/revisión), cliente/tienda y estado BORRADOR u OBSERVADO (actualizado 06/11/2026). Asigna numeración interna y número de revisión en servidor, admite borradores sin partidas y precios pendientes, y conserva snapshots de cliente/tienda/empresa. NuevaRevision permite crear un nuevo BORRADOR desde la última revisión enviada/observada/aprobada/rechazada/vencida, manteniendo la original; impide edición histórica y cambios cuando existe una OC. Cliente/tienda/tipo no cambian después de un envío o guía.

Partidas: admite TARIFARIO o MANUAL. El precio del tarifario, moneda y tratamiento se recuperan del servidor; se ignora un precio enviado por navegador para origen TARIFARIO. La edición conserva las referencias y descripciones anteriores aun con material desactivado. Un precio manual exige motivo y moneda igual al presupuesto. El cambio de moneda requiere confirmación explícita de revalorización y recalcula con precios en esa moneda, sin tipo de cambio. Moneda USD sin precio sigue NULL. Cantidades, precios convertidos e importes usan dos decimales; SQL recalcula y suma partidas sin aceptar totales externos. El redondeo es el mismo que BudgetAmounts.

CambiarEstado valida transiciones (actualizado 06/11/2026): BORRADOR (Registrado) y OBSERVADO a ENVIADO/CANCELADO; ENVIADO a OBSERVADO/APROBADO (Aceptada)/RECHAZADO (Rechazada)/CANCELADO. APROBADO/RECHAZADO pueden cancelarse si no hay OC ni guía vigente. VENCIDO se retira de filtros y destinos; solo conserva lectura histórica y posibilidad de cancelar registros antiguos. OBSERVADO admite corrección sobre la misma revisión y permanece observado hasta reenviarse. El formulario pide motivo exclusivamente al observar/rechazar/cancelar; MVC asigna fecha de Perú y usuario autenticado como responsable, manteniendo ambos parámetros obligatorios del SP. Envía solo con partidas e importes completos y datos de HYT presentes; aprobación guarda FechaAprobacion. Actualiza rowversions y audita fecha del evento, actor, responsable, observación y estados. La fecha del evento se presenta como fecha documental, sin conversión UTC que la desplace de día; los timestamps de registro conservan UTC.

IBudgetRepository y SqlBudgetRepository consumen exclusivamente SP para listados, fichas históricas, ayudas, guardado y transiciones. Registrados en DI; pantallas de presupuesto todavía pendientes. Nueva herramienta Hyt.DatabaseTests registrada en la solución: instala las fuentes en una base temporal local con nombre GUID, prueba operaciones reales y elimina solo esa base temporal. Las 24 comprobaciones SQL aprobadas incluyen snapshots, concurrencia, transición, precios pendientes, moneda, referencias desactivadas y mapeo ADO.NET. No deja usuarios, presupuestos ni datos comerciales de prueba en SistemaWebHYT.

### Pantallas de presupuesto y sugerencias — avance 03/11/2026

Aplicado `spu_Presupuesto_DatosIniciales`: consulta tasa vigente por fecha, datos únicos HYT y sugerencias de cliente, tienda, zona, centro FICO y proveedor/contrato activo más reciente. Verifica la pertenencia de la tienda al cliente. La API conserva las propiedades de snapshot en su JSON y no escribe datos. La ayuda tarifaria utiliza el código de tipo de servicio (longitud compatible con el snapshot de partida).

La etapa 5 parcial registra ocho fuentes; acumulado 117, con seis SP de presupuesto. Se reforzó Guardar para que un precio manual tampoco pueda enlazar una tarifa privada de otro cliente. Todos los cambios SQL de este avance se aplicaron a localhost\SQLEXPRESS / SistemaWebHYT.

BudgetController y vistas Index/Form/_Line conectan el módulo Presupuestos al menú. Listado con último mes inicial y conversión de límites inclusivos/exclusivos de Perú a UTC; búsqueda, página de veinte y retorno conservado. Captura incluye los campos de ambos Excel, partidas manuales/tarifario, datos sugeridos, revisión explícita de moneda y previsualización decimal calculada en servidor (sin persistencia). Al guardar, SQL verifica precios de origen y recalcula. Fichas de consulta deshabilitadas, historial paginado y acciones de estado con fecha/responsable/motivo. La nueva revisión se prepara desde la anterior y solo se crea al guardar.

Pruebas: 88 comprobaciones C#/HTTP y 26 de SQL real; Chrome verificó cliente/tienda/tarifario con zonas, USD sin precio, cambio a manual, cálculo, guardado MVC y consulta posterior, en escritorio/móvil sin desbordamiento. Capturas en entregables/preview/presupuesto-*.png. Se corrigió el registro duplicado de listeners en hyt-ui.js cuando se importa con y sin URL de versión, que impedía el guardado con protección de doble envío; ahora usa un estado compartido e inicialización única. Las pruebas de ayudas de mantenimientos siguieron pasando. Todavía no están implementados los dos reportes HTML ni la configuración editable de datos HYT ni los flujos de etapa 6.

### Configuración corporativa e impresiones HTML — avance 03/11/2026

Aplicados `spu_ConfiguracionHYT_Obtener/Guardar` para la fila única Id=1: permisos CONFIGURACION, RUC de once dígitos, longitudes, rowversion, responsables y auditoría paginada. El menú Configuración HYT abre esta ficha y ofrece acceso a vigencias IGV. No reescribe snapshots de revisiones anteriores.

Script 014 amplía PresupuestoRevision con SimboloMoneda (desde la tabla Moneda) y TiendaCodigo, junto a los campos de snapshot ya documentados. Guardar conserva código de tienda y completa el centro FICO desde la tienda cuando la partida no lo especifica. Todos los campos nuevos son idempotentes y se aplicaron directamente a la base local. No se modificaron los Excel originales ni se insertaron datos corporativos de prueba en la base de trabajo.

`spu_Presupuesto_RegistrarImpresion` recibe HTML generado exclusivamente en el servidor, formato ANTERIOR/NUEVO, versión de plantilla y rowversion de revisión. Verifica PRESUPUESTOS.CONSULTAR y REPORTES.CONSULTAR, detecta cambios concurrentes, conserva el contenido y SHA256 de sus bytes UTF-16 LE y audita la generación en la revisión. `ObtenerImpresion` recupera el contenido original; el repositorio comprueba su huella antes de entregarlo. `ListarImpresiones` devuelve veinte registros y total. Una nueva impresión no reemplaza generaciones anteriores ni cambia el estado comercial. La descarga requiere los mismos permisos.

BudgetPrintRenderer genera HTML autónomo con logo PNG y CSS incluidos, datos del snapshot y precios/importes de la revisión. El contenido se sirve con CSP sin scripts ni recursos externos y puede imprimirse desde el navegador. Formato anterior: siete columnas, parciales y subtotal de grupo; formato nuevo: once columnas, referencias de tienda/FICO/necesidad/tipo/posición/SAP. Los totales globales provienen de los importes persistidos, sin sumar grupos por segunda vez. Ambos muestran moneda real y tratamiento/tasa IGV parametrizados, cantidades y precios con dos decimales. BORRADOR/precios pendientes se identifican; no se inventan importes.

Verificación: 103 comprobaciones HTTP/C#, 37 SQL reales en base temporal eliminada, siete casos de impresión Chrome/PDF. Casos simples de ambos formatos, vacío y USD: una página A4. Cuarenta partidas: 10 páginas en formato nuevo y 6 en anterior, con todas las filas/encabezados y total 4,720.00. Descripción extensa: texto final conservado, sin caracteres fuera de página. Capturas renderizadas revisadas visualmente. Los PDF son evidencia de prueba; el producto genera HTML. Se contrastó contenido/orden con las dos referencias Excel y se corrigió el pie huérfano del formato nuevo. Fuentes acumuladas 122; etapa 5 parcial 13 (once SP y dos scripts). La etapa 6 todavía está pendiente.

## Persistencia inicial de seguimiento — 03/11/2026

Etapa 6 en curso. Aplicadas directamente en localhost\SQLEXPRESS / SistemaWebHYT las quince fuentes del manifiesto etapa 6: tabla DocumentoOperacion, script 016, doce SP Seguimiento (Guardar/Obtener/Listar para OC, GUIA, CONFIRMACION y EJECUCION) y permisos 017. Acumulado: 137 fuentes.

DocumentoOperacion conserva la operación GUID, actor, tipo, padre, registro, archivo y huella de solicitud para reintentos idempotentes. Guardados verifican permisos, versiones de padre/documento, reglas de emergencia/invitación y cardinalidad. Adjuntos se registran con metadatos y auditoría en la misma transacción; almacenamiento físico privado y pantallas todavía pendientes. Importes aceptan máximo dos decimales.

Pruebas actuales: 53 comprobaciones de SQL real en base temporal eliminada. Incluyen guía de emergencia sin OC, números alfanuméricos, reintento sin duplicar datos/auditoría, OC de invitación solo tras aprobación, una OC por presupuesto y una confirmación vigente por OC. Obtener/Listar de guía, OC y confirmación ejecutados con datos. Se corrigió SELECT INTO de listados convirtiendo explícitamente ambos rowversion a BINARY(8): SQL no permite dos columnas timestamp en la tabla temporal. La corrección fue reaplicada a la base local.

Pendiente: transiciones documentales, integración MVC de seguimiento, adjuntos privados consultables, pruebas completas de ejecución/archivos y conciliación de requisitos. Etapas 5/6 no declaradas terminadas. La base de trabajo no recibió registros de prueba.
## Integración inicial MVC de seguimiento — 03/11/2026

Añadidos IDocumentRepository/SqlDocumentRepository con catálogo cerrado de tipos y operaciones, consultas y guardado ADO.NET exclusivamente por SP. Registrado en la aplicación. La selección de procedimiento no admite nombres arbitrarios; errores de negocio y disponibilidad tienen tratamiento controlado.

FollowupController ofrece /Seguimiento/oc-cliente, /guias, /confirmaciones y /ejecucion con permiso específico, último mes por defecto, límites UTC desde fechas Perú, búsqueda y página SQL de veinte. Fichas consultan contexto y auditoría paginada; importes con dos decimales y fechas de evento sin cambiar día. El menú de OC y guía redirige al listado operativo. Esta entrega incorpora consulta; todavía no incorpora formularios de registro/edición ni acciones popup, transiciones o visualización de archivos.

Verificación: compilación sin advertencias/errores; 58 comprobaciones SQL reales en base temporal eliminada, incluidas ficha/auditoría ADO.NET, filtros con/sin confirmación, guardado de ejecución pendiente y rechazo por permisos. Pruebas HTTP existentes y diez nuevas de seguimiento aprobadas. Chrome revisó cinco rutas (listados OC/guía y fichas OC/guía/confirmación) en escritorio y móvil sin desbordamiento. Diez capturas en entregables/preview/seguimiento-*.png con datos ficticios; ficha móvil revisada visualmente y espaciado corregido. Host temporal de pruebas cerrado y archivo temporal de credenciales ficticias eliminado.

Siguen pendientes captura/edición documental, botones popup en presupuestos/OC, adjuntos privados consultables, transiciones de estado, pruebas integrales y paquetes finales. Objetivo 5/6 activo; no se ha aprovisionado administrador ni conexión web real.
## Captura emergente y adjuntos privados — 03/11/2026

Añadidos DocumentFields (catálogo cerrado y validación), formularios MVC de registro/edición y guardado por IDocumentRepository. Se validan campos obligatorios, fechas ISO sin hora, revisiones, motivo de edición, importes no negativos con hasta dos decimales y tokens de ocho bytes. El identificador de operación se conserva al fallar; no se sustituyen versiones obsoletas por versiones actuales durante un reintento. El contenido y el estado comercial no se aceptan como JSON libre del navegador.

Grilla compartida admite acciones adicionales. Presupuestos ofrece guía/acta, OC y ejecución conforme a permisos; OC ofrece confirmación del cliente. Captura en dialog con estilo HYT, calendario FralseCont, foco, carga, errores, bloqueo de doble envío y retorno local conservando filtros. El guardado parcial devuelve 422 cuando hay errores; al completar vuelve al listado. Acciones de grilla separadas en áreas de clic de al menos 44 px. Calendario ahora limpia paneles y eventos de cada ámbito antes de reinicializar formularios dinámicos.

Adjuntos: guía solo PDF; OC/confirmación PDF, PNG o JPG; máximo 15 MB. El servidor comprueba extensión, firma inicial, tamaño recibido, nombre limpio y SHA256. Clave opaca GUID de operación con extensión normalizada. Escritura a temporal propio y traslado atómico sin sobrescritura; reintento idéntico reutiliza el archivo. Archivo privado fuera de wwwroot, configurable con DocumentStorage:Root; por defecto LocalApplicationData/SistemaWebHYT/Adjuntos. En despliegue, configurar almacenamiento persistente y sus permisos. Ante resultado SQL incierto se conserva el archivo para no eliminar una evidencia posiblemente confirmada; los archivos huérfanos requieren conciliación administrativa, sin limpieza automática destructiva.

Aplicado directamente en SistemaWebHYT local spu_Seguimiento_ObtenerArchivo y ampliación de permisos 017. Consulta exige actor activo y permiso del tipo documental, vínculo correspondiente y archivo no retirado. Ruta MVC /Seguimiento/{tipo}/Archivos/{id}: visualizar o descargar, sin exponer clave física, con nosniff, política privada sin caché y CSP. Antes de entregar se verifican tamaño y huella. Retiro por selección en edición exige motivo, conserva metadatos y auditoría, y deja de permitir consulta. Las fichas muestran los archivos y sus enlaces.

Verificación: 64 comprobaciones SQL reales aprobadas en base temporal eliminada. Incluyen subida lógica, nombre Unicode, autorización, vínculo entre tipo y archivo, reintento sin duplicar metadatos/auditoría y retiro auditado. Pruebas HTTP/C# agregadas para captura popup, campos permitidos, antifalsificación, motivo/versiones, dos decimales, carga multipart real al almacenamiento temporal, visualización/descarga y permisos. Pruebas de archivos físicos: nombre limpio, clave opaca, hash, reintento, rechazo de falso PDF/imagen en guía, contenido modificado y salida de carpeta. Carpetas temporales eliminadas con validación de ruta.

Chrome: registro guía desde presupuesto, confirmación desde OC, guardado y retorno, tres aperturas OC sin calendarios duplicados, carga de PDF desde selector de archivo y recuperación privada. PDF visualizado y captura revisada; popup móvil y escritorio revisados sin desbordamiento. Evidencias seguimiento-popup-guia-*.png y seguimiento-pdf-privado.png en entregables/preview; son datos ficticios. El host temporal de pruebas quedó cerrado. Acumulado: 138 fuentes SQL; manifiesto etapa 6: 16 fuentes.

Pendientes principales: transiciones de estado de OC/guía/confirmación/ejecución con auditoría e idempotencia, regularización comercial de emergencia con OC, trazabilidad completa del expediente, pruebas integrales y paquetes finales. El objetivo 5/6 sigue activo y no se declara terminado.
## Estados documentales — 03/11/2026

Aplicados directamente spu_SeguimientoOC_CambiarEstado, spu_SeguimientoGUIA_CambiarEstado, spu_SeguimientoCONFIRMACION_CambiarEstado y spu_SeguimientoEJECUCION_CambiarEstado, más permisos 017. Generador: tools/prepare_document_states.py; fuentes separadas por SP. Manifiesto etapa 6: 20 fuentes; acumulado: 142 scripts, 47 tablas y 105 relaciones validados offline.

Cada transición administra una transacción propia, exige permiso de edición, tokens de documento/padre, operación GUID, fecha y responsable; registra estados anterior/nuevo, motivo y snapshots en EventoAuditoria. La misma operación idéntica devuelve el resultado sin duplicar auditoría; reutilizarla con otros datos se rechaza. Bloquea modificaciones del expediente marcado facturado y respeta orden de bloqueo operación → presupuesto → OC/documento. Motivo obligatorio en observación, anulación, cancelación, retiro y reaperturas.

Guía: registrada/enviada/observada puede avanzar a aceptación; al aceptar se conservan fecha de evento y representante. Puede observarse o anularse con motivo; una OC cerrada obliga a reabrir antes de revocar su guía aceptada. Al retirar aceptación se limpian campos de aceptación actuales, conservando valores anteriores en auditoría.

OC: registrada → trabajo completado exige ejecución terminada; trabajo completado → cerrada operativamente exige además guía aceptada. Puede reabrirse a trabajo completado y luego registrada. Cancelar exige retirar confirmación vigente, anular guía y detener ejecución activa. La reactivación conserva el mismo registro único por presupuesto.

Confirmación: vigente → retirada con motivo; no se reactiva el registro histórico. Puede registrarse otra confirmación, manteniendo la restricción de una vigente y conservando las anteriores.

Ejecución: pendiente → programada exige fecha programada y responsable del trabajo. Inicio exige responsable; invitación necesita OC y programación, emergencia permite iniciar sin OC. Término requiere fecha no anterior al inicio. Reabrir trabajo terminado exige reabrir OC que figure completada/cerrada y revisar guía aceptada; se limpia término actual sin borrar antecedentes.

Añadido ChangeStateAsync al repositorio ADO.NET y acciones/formulario de estados en fichas MVC, con permiso, antifalsificación, fecha, responsable, motivo y versiones. La UI muestra las transiciones propias del estado; SQL conserva la autoridad sobre dependencias. No cambia automáticamente la aprobación comercial del presupuesto.

Verificación actual: 79 comprobaciones SQL reales aprobadas en base temporal eliminada. Nuevos casos cubren orden de estados, fechas, emergencia sin OC, aceptación idempotente, versiones anteriores, trabajo completado/cierre y reaperturas bloqueadas por dependencias, retiro y reemplazo de confirmación con historia, ADO.NET y permisos. Pruebas HTTP de formulario/acción, fecha obligatoria y edición autorizada aprobadas. Compilación sin advertencias ni errores.

Pendiente para cierre 5/6: regularización comercial de emergencia después de OC, trazabilidad consolidada del expediente y actualización visible del presupuesto tras documentos; ampliar pruebas de bordes y de las acciones en navegador, conciliar requisitos y preparar paquetes finales. Objetivo activo.
## Decisión vigente: Cloudflare R2 y Turnstile — 03/11/2026

Por instrucción del usuario, los adjuntos de la aplicación se guardan en Cloudflare R2. El administrador creará luego el bucket `hyt-documentos`, declarado en `.env.example` como CloudflareR2__BucketName y en el valor predeterminado de CloudflareR2Settings. Esta decisión reemplaza el almacenamiento local descrito en avances anteriores; el servicio local se conserva solo como sustitución en pruebas aisladas. No se creó ningún bucket ni se guardaron claves reales.

R2DocumentFileStore es el proveedor registrado en DI. CloudflareR2ObjectClient utiliza SDK S3, endpoint derivado de AccountId, credenciales explícitas, región auto, HTTPS, escritura condicional If-None-Match y opciones de compatibilidad R2. Validación común DocumentUploadRules conserva restricciones y SHA256. Reintento no sobrescribe contenido, lectura verifica integridad y permisos SQL/MVC siguen vigentes. No hay fallback a almacenamiento local ni enlaces públicos. No se requiere cambio de tablas o SP para el proveedor: la clave de objeto conserva el formato GUID/extensión ya validado por SQL.

Se integró Turnstile en login: widget público con acción login y validación Siteverify del lado servidor antes de acceder al repositorio de usuarios. Verifica éxito, acción, hostname permitido y vigencia. Falla cerrada ante token ausente, inválido o error de red. En producción es obligatorio; las pruebas lo desactivan explícitamente en Development o simulan el servicio. Claves exclusivamente por entorno, secret key ausente del HTML. Configuración y preparación en CONFIGURACION_CLOUDFLARE.md; variables sin valores secretos en .env.example. Los archivos .env privados están ignorados.

Pruebas nuevas de R2: conservación de contenido, huella, reintento sin duplicación/sobrescritura y rechazo de objetos alterados. Turnstile: éxito, acción y dominio ajenos, expiración, token ausente/largo, rechazo del servicio, JSON inválido, red y obligatoriedad en producción. Pruebas HTTP de widget sin secreto y bloqueo antes de consultar usuario al rechazar un token; éxito de Siteverify permite continuar a contraseña. No son pruebas contra un bucket real, pendiente de creación/configuración por el usuario conforme a su instrucción.

Avance previo de este tramo: regularización comercial de emergencia con OC aplicada en SQL (solo enviar/aprobar revisión vinculada, sin editar contenido/importes); historial consolidado del expediente paginado y filtrado por permisos; última actualización documental visible en listado de presupuestos. 90 comprobaciones SQL reales aprobadas, incluidas reversión completa ante fallo de auditoría, regularización sin recalcular, historial repetido de más de veinte eventos y protección de documentos sin permiso. OC muestra moneda/importe/confirmación y filtro con/sin confirmación. Continúan pendientes conciliación final de etapas 5/6, pruebas de navegador de los últimos ajustes y paquetes finales.
## Cierre de implementación y verificación de etapas 5/6 — 03/11/2026

Presupuestos y seguimiento documental conciliados contra el plan en VERIFICACION_ETAPAS_5_Y_6.md; recorridos en GUIA_ETAPAS_5_Y_6.md. Paquetes ETAPA_5_HYT.sql y ETAPA_6_HYT.sql y SHA256_ETAPA_5/6.json preparados en entregables. Instalación desde paquetes aprobada con 97 comprobaciones SQL reales; 164 comprobaciones HTTP/C# aprobadas. Chrome verificó presupuestos, seguimiento, ayudas, archivos, estados, regularización, filtro de OC confirmadas y presentación móvil/escritorio. Siete casos de impresión aprobados y contenido dentro de A4, con encabezados repetidos y todos los renglones conservados.

Validación offline: 142 fuentes, 47 tablas y 105 relaciones, más cinco pruebas de regresión del validador. Base local: 47 tablas, 78 SP y cero presupuestos/adjuntos de prueba. Los 28 SP de etapas 5/6 se compararon con sus definiciones instaladas en SistemaWebHYT y coinciden con las fuentes. Compilación sin advertencias ni errores; host y credenciales ficticias de prueba cerrados/eliminados.

Presupuesto_Obtener agrega EstadoOC y RevisionOC para regularización y consolida historial de presupuesto/revisión seleccionada, OC, guía, ejecución, confirmaciones y archivos, limitado por permisos específicos. Presupuesto_Listar toma actualización y responsable del presupuesto raíz, por lo que refleja cambios documentales. Se conserva paginación de veinte y conteo coherente. UI identifica la entidad con etiqueta legible y ofrece navegación al expediente completo desde popups.

User Secrets propio ya contiene la conexión SQL local, con autenticación de Windows, y campos Cloudflare. R2 y Turnstile están implementados y probados con sustitutos; el usuario creará el bucket hyt-documentos y completará las claves después, según su instrucción. La prueba real contra Cloudflare no se declara realizada. Aprovisionar el primer administrador según la guía de acceso; no hay credenciales predeterminadas. Marcado Facturado, dashboard y publicación permanecen en etapas 7/8/9.
## Etapa 7 iniciada — 03/11/2026

Aplicados directamente en SistemaWebHYT local spu_Facturado_Guardar, spu_Facturado_Obtener y Script/018_PermisosEtapa7.sql. No se crearon tablas nuevas: se utiliza ReferenciaComprobanteExterno y DocumentoOperacion. Generador tools/prepare_invoice_marks.py; manifiesto incremental etapa7 con tres fuentes; acumulado 145 fuentes.

Guardar registra fecha, código textual 01/03, serie y número alfanuméricos, vinculados a OC y confirmación vigente. Mantiene los requisitos documentales existentes del trigger 007: OC válida, ejecución terminada y guía aceptada. No emite comprobantes ni asientos ni se comunica con SUNAT. Corrección y retiro requieren motivo y versiones; retiro conserva referencia anterior. Auditoría y actualización de OC/presupuesto son atómicas. Operación GUID con huella evita duplicados y rechaza reutilización con otros datos. Obtener devuelve contexto, referencia vigente e historial de veinte referencias por página.

Validación: 102 comprobaciones SQL reales aprobadas en base temporal eliminada. Cinco nuevos casos prueban bloqueo sin confirmación, código 01 y ceros conservados, reintento idempotente, una referencia vigente y retiro con motivo/auditoría. Fuente corregida y aplicada para validar requisitos antes de que actúe el trigger de protección existente. Sin datos comerciales de prueba en la base de trabajo.

Pendiente del objetivo 7/8/9: repositorio y popup/listado de marcado externo, auditoría visible y pruebas web; dashboard y reportes HTML conciliados por moneda; validación integral y preparación de entrega/despliegue. Bucket R2 y claves Cloudflare siguen pendientes del usuario conforme a sus instrucciones. Objetivo activo, sin declaración de cierre.
## Integración MVC del marcado externo — 03/11/2026

IInvoiceMarkRepository y SqlInvoiceMarkRepository consumen exclusivamente spu_Facturado_Obtener/Guardar. InvoiceMarkController expone /OCCliente/{id}/Facturado con permisos FACTURADO y OC_CLIENTE, antifalsificación, tokens de concurrencia y GUID de operación. El listado OC ofrece botón popup Facturado; reutiliza calendario, ventana y control de doble envío existentes. Guarda fecha, código 01/03, serie y número; permite corrección y retiro con motivo. Muestra historial de referencias vigente/retiradas, con página SQL de veinte. No emite comprobantes ni los envía a SUNAT.

Pruebas HTTP nuevas aprobadas: parcial popup, campos/códigos, conservación de ceros, retorno local seguro, rechazo de registro incompleto y consulta sin edición. Pruebas SQL ampliadas con repositorio real para lectura de referencia retirada y marcado Boleta código 03. Compilación sin errores ni advertencias. Pendientes de etapa 7: exponer referencia y estado en grilla/ficha OC, auditoría detallada visible, pruebas de retiro/corrección en navegador y conciliación final. Etapas 8 y 9 continúan pendientes; objetivo activo.
## Visibilidad de referencia y auditoría — 03/11/2026

spu_Facturado_Obtener incorpora auditoría detallada y total de eventos, en páginas de veinte, con usuario, fechas, motivo y JSON anterior/nuevo. El formulario MVC muestra las acciones y los cambios conservados. La ficha OC enlaza directamente a Facturado.

Los SP Seguimiento_Listar incorporan un objeto Facturado con fecha, tipo, serie, número, fecha/usuario de marcado. El objeto se devuelve únicamente al tener FACTURADO.CONSULTAR; otros permisos de OC no revelan estos datos. Grilla MVC muestra columnas adicionales únicamente con ese permiso. Fuentes modificadas aplicadas directamente en SQL local, conservando manifiestos de etapas posteriores al regenerar seguimiento.

Verificación: 107 comprobaciones SQL reales aprobadas, incluidas corrección con valor anterior consultable, referencia actual en listado y ocultación ante falta de permiso específico. Pruebas MVC aprobadas y Chrome verificó captura/guardado desde popup OC, recorrido documental previo y escritorio/móvil. Host temporal cerrado. Compilación sin advertencias/errores. Pendiente: ampliar casos de retiro/corrección de UI y consolidar paquetes de etapa7; comenzar dashboard/reportes etapa8 y preparación/validación etapa9. Objetivo activo.
## Etapa 8: consultas de indicadores y reportes — 03/11/2026

Aplicados directamente spu_Dashboard_Obtener, spu_Reporte_Listar y Script/019_PermisosEtapa8.sql. Generador prepare_reporting.py y manifiesto etapa8 con tres fuentes. No se crearon tablas; acumulado 148 fuentes, 47 tablas y 105 relaciones. Generadores de etapas anteriores conservan los manifiestos posteriores al regenerarse.

Dashboard usa cohorte por fecha de registro de presupuesto/OC, con rango UTC enviado desde fechas Perú. Cada presupuesto cuenta una sola vez con su última revisión; cada OC una vez con su revisión vinculada. Importes separados PEN/USD; presupuestos sin precio se contabilizan explícitamente. Indicadores: estados, sin confirmar, confirmadas sin marcado, listas para marcar (confirmación + ejecución terminada + guía aceptada), facturadas y movimientos diarios agrupados en fecha Perú. Requiere REPORTES.CONSULTAR y limita cada fuente por permiso de presupuestos/OC. No mezcla monedas ni suma categorías superpuestas.

Reporte_Listar ofrece GANADOS (última revisión aprobada por presupuesto, excluido expediente cancelado), OC_COMPLETADAS (ejecución terminada), OC_CON_ACTA (guía no anulada) y OC_PENDIENTES_FACTURA (OC no cancelada sin referencia vigente). Categorías superpuestas, no sumables. Cohorte por registro, veinte filas SQL por página, conteo y totales de toda la consulta por moneda. Pendientes distinguen sin confirmar, trabajo pendiente, guía pendiente y lista para marcar.

Pruebas reales: 117 comprobaciones aprobadas en base temporal eliminada; nuevas verifican conteo sin duplicar revisiones/confirmaciones históricas, separación PEN/USD, paginación sin duplicados de cuatro reportes, exclusión de OC facturada, permisos y cambio de día UTC/Perú. Validación offline aprobada. Pendientes: repositorios y dashboard MVC con gráficos profesionales/interacción, reportes HTML y conciliación visual; cierre de pruebas/paquetes etapa7; preparación y validación etapa9. Objetivo activo.
## Reportes MVC y decisión de publicación — 03/11/2026

Implementados IReportingRepository/SqlReportingRepository por ADO.NET con SP; ReportingDates convierte límites inclusivos/exclusivos desde fechas Perú a UTC. ReportsController y /Reportes ofrecen catálogo cerrado de cuatro reportes, permisos REPORTES y documento fuente, rango inicial último mes, veinte filas y totales de toda la consulta por moneda. HTML codifica datos y enlaza presupuesto/OC; imprime la página consultada en A4 horizontal con encabezado de tabla repetido. Menú Reportes redirige a esta vista. No hay consultas SQL en Razor.

Pruebas HTTP aprobadas para cuatro vistas, paginación, ceros/códigos, importes, codificación de datos y permisos. Pruebas con repositorios SQL reales incorporadas para filas/conteos/totales y dashboard autorizado. Pendientes de etapa8: dashboard visual del inicio con curvas, barras/dona, interacciones, vistas rápidas y revisión visual/impresión de reportes. Etapa7 conserva pendientes de cierre de casos UI/paquetes.

Instrucción expresa del usuario para etapa9: dejar la publicación preparada; no publicar todavía porque faltan claves y pruebas manuales. No solicitar nuevamente dominio como bloqueo para la preparación: los artefactos de despliegue se dejarán parametrizados. Bucket R2 hyt-documentos y claves continúan pendientes del usuario. Objetivo 7/8/9 activo.
## Dashboard visual del inicio — 03/11/2026

HomeController ahora consulta IReportingRepository solo con REPORTES.CONSULTAR y valida fechas Perú. Perfiles sin permiso conservan bienvenida y módulos autorizados, sin consultar ni recibir indicadores. DashboardModel/_Dashboard muestra cantidades por fuente autorizada, OC sin confirmar, listas para marcar, importes separados PEN/USD y enlaces a listados con fechas conservadas. No mezcla documentos como si fueran servicios distintos.

Gráficos SVG nativos: curvas de registros diarios con interpolación monótona entre puntos, columnas por estado y dona de tres categorías disjuntas de OC activas. Selección de series, foco y títulos de puntos, selector de intervalo utilizable en móvil y detalle tabular accesible. Rangos largos se agrupan en hasta noventa intervalos para acotar recursos. Días sin registros se incluyen con cero; sin series o datos se muestra estado vacío. Colores HYT rojo/negro/gris y CSS adaptable; sin dependencia de librerías externas ni animaciones obligatorias.

Pruebas HTTP nuevas aprobadas: inicio conectado, importes por moneda, tres gráficos y detalle accesible, rechazo de fechas invertidas y ausencia de consulta del repositorio para perfiles sin reportes. Chrome verificó dos curvas, ocultar/restaurar series, estado sin selección, detalle por fecha coincidente, barras/dona, rangos largos acotados y escritorio/móvil sin desbordamiento. Capturas dashboard-escritorio.png y dashboard-movil.png en entregables/preview revisadas visualmente; cifras de esas capturas son exclusivamente de repositorio de pruebas. Se mejoraron etiquetas móviles y geometría de dona tras revisión. Host temporal cerrado.

Pendientes del objetivo: casos adicionales de datos vacíos/parciales y permisos de fuentes, validación visual/impresión de reportes, cierre de pruebas/paquetes etapa7/8 y preparación parametrizada de publicación etapa9. El usuario indicó expresamente no desplegar aún: faltan claves y pruebas manuales. Objetivo activo.
## Preparación de etapa 9, sin publicación — 03/11/2026

Añadidos Dockerfile/.dockerignore, deploy/compose.hyt.yml, plantillas de parámetros privados/no secretos y Nginx. Servicio independiente en loopback, redes externas verificables para SQL/salida, volumen persistente de claves y ejecución no root. Dominio/IP/redes son parámetros pendientes. Guía GUIA_PUBLICACION_Y_QA.md incluye preparación, aceptación manual, respaldo SQL/objetos/claves, restauración aislada y reversión de versión.

Program configura encabezados reenviados solo desde IP explícita de proxy, con límite de una etapa y simetría IP/protocolo; exige IP válida en producción. Middleware anterior a redirección HTTPS. GET /health/live público devuelve solo ok como sonda de vida, sin comprobar ni exponer SQL/Cloudflare. Producción conserva DataProtection persistente requerido y Turnstile obligatorio.

Verificado: dotnet publish Release aprobado, sin archivos secrets.json/.env/certificados privados en salida; Docker Compose config aprobado sin resolver secretos; pruebas HTTP/C# de sonda y confianza limitada de proxy aprobadas. Motor Docker local no disponible: no se construyó ni ejecutó imagen y no se inició Docker Desktop. Activación, claves, bucket y pruebas manuales siguen diferidas expresamente por el usuario. Objetivo continúa activo: cerrar pruebas y paquetes 7/8 y revisar preparación completa de 9.

## Dashboard, navegación y consultas vacías — 05/11/2026

spu_Dashboard_Obtener agrega parámetros opcionales @Cliente/@Tienda NVARCHAR(200), @PaginaPendientes INT=1. Filtra nombres históricos de cliente y tienda en todas las fuentes. ActualizadoUTC identifica la consulta. Devuelve VerPendientes, PendientesPeriodo, TotalPendientes y Pendientes con paginación SQL de veinte y antigüedad, incluyendo expedientes anteriores al rango hasta HastaExclusiva. La bandeja requiere PRESUPUESTOS, OC_CLIENTE y GUIAS.CONSULTAR. Los pendientes pueden solaparse por causa; cada expediente se lista una vez.

Movimientos agrega Confirmaciones y Marcados: MIN(FechaRegistro) por OC y clase, filtrado por fecha de ese primer registro. Una corrección, retiro o nuevo registro posterior no incrementa el primer evento. Las series de presupuestos y OC conservan su propia fecha de registro. No se interpreta el gráfico como ingresos ni como documentos emitidos. Las listas JSON vacías del dashboard se normalizan a [] para conservar el contrato.

spu_Reporte_Listar, spu_Presupuesto_Listar y los cuatro spu_Seguimiento*_Listar agregan @Cliente/@Tienda opcionales, antes de conteo, paginación e importes. Los accesos del dashboard conservan estos parámetros y estado. Reporte_Listar devuelve [] si no hay filas; el repositorio también acepta NULL defensivamente. No se altera ninguna tabla. Fuentes y firmas de 03/11 se conservan; se añaden firmas del 05/11 en los SP modificados. SQL local actualizado y 82 definiciones comparadas con fuentes. Ver VERIFICACION_ETAPAS_7_8_Y_9.md.

spu_Reporte_Listar incorpora @Situacion NVARCHAR(30)=NULL: SIN_CONFIRMAR, CONFIRMADAS, LISTAS o FACTURADAS. Filtra antes de conteo, totales y página; rechaza seguimiento OC en GANADOS. Los cuatro accesos rápidos del dashboard se conciliaron exactamente contra sus conteos en pruebas SQL (136 comprobaciones totales). Formularios/paginación/impresión identifican el seguimiento elegido.

## Respaldo de acceso y espera común — 05/11/2026

ManualCaptchaService y AccessController permiten verificar un captcha de un intento/cinco minutos ante falta o fallo de Turnstile, conservando autenticación por los SP existentes. hyt-loading.js centraliza la espera y bloqueo de doble envío para las vistas MVC, AJAX y descargas. No hay cambios de estructura SQL, SP, permisos o credenciales. Documentación funcional y verificación en CAPTCHA_Y_LOADING.md.

## Códigos internos y tarifarios desde materiales — 05/11/2026

Script/020_CodigosInternosAutomaticos.sql crea de forma idempotente dbo.SecuenciaCodigoInterno (BIGINT, sin ciclo) y el índice único UX_ServicioMaterial_CodigoInterno. No renumera registros existentes. Los Guardar de Cliente, Tienda, Zona, Especialidad, TipoServicio, ServicioMaterial y Tarifario asignan prefijos CLI, TIE, ZON, ESP, TS, MAT y TAR; secuencia compartida, relleno mínimo de seis dígitos y comprobación de colisiones con códigos históricos. El bloqueo transaccional por entidad evita duplicados en altas simultáneas. Los saltos de numeración son normales. Al editar, recuperan el código original y descartan intentos de cambiarlo desde JSON.

UnidadMedida mantiene Id como código interno automático y Codigo como abreviatura editable y única; Guardar informa duplicados de abreviatura. Listar busca también por Id. ServicioMaterial.Posicion se conserva físicamente para compatibilidad histórica, pero deja de formar parte del maestro, búsqueda, JSON Obtener y guardado. Posicion de TarifarioPartida continúa disponible. ProveedorCliente_Listar expone la identidad interna como código de fila y mantiene SAP como dato externo. Los mantenimientos sin un código de negocio generado muestran el Id interno como solo lectura.

spu_MaterialTarifas_Listar recibe actor, material y página, exige MATERIALES.CONSULTAR y TARIFARIOS.CONSULTAR, valida argumentos y devuelve filas JSON más total en un segundo resultado. Pagina veinte en SQL con orden estable FechaRegistro/Id descendente. Incluye versión/estado, posición, SAP, tipo de servicio, PEN/USD, moneda, IGV, zonas y responsables/fechas. Concede ejecución al rol limitado hyt_app_executor; no añade permisos funcionales. Reutiliza tablas existentes, sin duplicar precios en ServicioMaterial.

MVC permite alta/edición de partidas desde una ventana contextual en la ficha del material y crear tarifario → versión → precio en el mismo diálogo. Valida que el material coincida en GET y POST. Cada paso del recorrido se guarda independientemente con los SP y auditoría existentes. Publicados siguen inmutables; su versión puede copiarse para obtener un borrador. El listado se refresca conservando los campos aún no guardados del material. Los menús Especialidades y Tipos de servicio son independientes y reutilizan los permisos de materiales.

Aplicación local completada sin datos de prueba comerciales: script 020, doce SP de mantenimiento y nueva consulta de precios. Comparadas las 83 definiciones instaladas con fuentes. Manifiestos de etapas 3/4 y acumulado actualizados: 150 fuentes, 47 tablas, 105 relaciones; conservadas etapas posteriores. Ver CODIGOS_Y_TARIFAS_DESDE_MATERIALES.md para funcionamiento y evidencia: 149 comprobaciones SQL, 220 HTTP/C#, navegador escritorio/móvil y publicación Release preparada.

## Identidad del cliente por tipo de persona — 05/11/2026

spu_MantCliente_Guardar exige tipo y número de documento, además del nombre correspondiente a TipoEntidad. JURIDICA admite RUC y razón social; NATURAL admite DNI o CE y nombres/apellidos. Rechaza combinaciones incompatibles y limpia Nombres/Apellidos para jurídica o RazonSocial para natural antes de guardar y generar el JSON de auditoría. Conserva permisos, código automático, control de concurrencia y transacción. No altera tablas ni migra datos históricos; al editar clientes antiguos incompletos deberá completarse su documento.

El catálogo cerrado IdentityDocument (RUC, DNI, CE) se sirve desde el repositorio de mantenimiento, sin mantenimiento adicional ni tabla nueva. TipoIdentificacion pasa de texto libre a selección. MVC y JavaScript habilitan los campos según la persona, señalan los obligatorios y cambian la etiqueta a Número de RUC/DNI/CE; los campos no aplicables se ven deshabilitados. MaintenanceInput valida la combinación y descarta sus valores no aplicables incluso ante envíos manipulados. La validación comprueba presencia y tipo compatible; no consulta RENIEC, Migraciones o SUNAT ni verifica la autenticidad del número.

Validación aprobada: 153 comprobaciones SQL en base temporal eliminada, pruebas HTTP/C# ampliadas para las combinaciones y limpieza de campos, navegador con cambios entre persona jurídica/natural y escritorio/móvil. Fuentes, paquete de etapa 3 y huellas actualizados. Procedimiento aplicado a SistemaWebHYT local sin modificar clientes existentes.

## Ajuste vigente: tipo de documento oculto — 05/11/2026

Por corrección del usuario, TipoIdentificacion ya no se presenta ni se exige o valida en Clientes. La definición lo marca Hidden, la vista no genera su control y MaintenanceInput ignora cualquier valor enviado. Se retiró el catálogo IdentityDocument. El número continúa obligatorio: etiqueta Número de RUC para jurídica y Número de DNI / CE para natural. Se conserva la habilitación de razón social o nombres/apellidos según TipoEntidad.

spu_MantCliente_Guardar no lee ni valida TipoIdentificacion del JSON. En ediciones conserva el valor histórico desde el registro bloqueado; nuevas altas lo dejan NULL. No se elimina la columna ni se modifican datos existentes. Se retiraron obligatoriedad, longitud y compatibilidad por tipo de persona para ese campo. Las restantes validaciones de número, nombre, permisos y concurrencia continúan. Esta decisión reemplaza la selección RUC/DNI/CE indicada en el apartado anterior.

Procedimiento actualizado en SQL local y paquetes de etapa 3 regenerados. Compilación sin errores/advertencias; pruebas HTTP y 153 comprobaciones SQL aisladas aprobadas, incluidas altas de ambas personas sin TipoIdentificacion.

## Código SAP en el maestro de Clientes — 05/11/2026

Cliente incorpora CodigoSAP NVARCHAR(60) NULL, opcional, alfanumérico y sin unicidad global: es un dato por cliente, independiente del código interno automático. Actualizados la definición inicial de tabla y el generador, y creado Script/021_CodigoSAPCliente.sql para instalaciones existentes. La migración añade la columna una sola vez y traslada el valor anterior de ProveedorCliente, priorizando activo, FechaRegistro e Id descendentes; no elimina las fuentes. No vuelve a poblar valores borrados por el usuario al repetir el script.

spu_MantCliente_Obtener/Guardar incluyen CodigoSAP y la auditoría conserva sus cambios junto con el cliente. Listar permite buscar por ese dato. Campo visible «Código SAP» dentro del formulario de Clientes, con longitud máxima de sesenta caracteres y conservación de ceros iniciales. El acceso «Datos de HYT por cliente» se excluye de los registros relacionados. Su listado anterior redirige a la ficha del cliente cuando se conoce parentId o al listado de Clientes. Las tablas y auditorías anteriores se conservan para consulta histórica.

spu_Presupuesto_DatosIniciales toma Cliente.CodigoSAP y lo devuelve bajo la clave compatible Proveedor.CodigoSAPProveedor. Funciona aunque no exista fila en ProveedorCliente. El contrato marco anterior sigue consultándose cuando corresponda. Las revisiones e impresiones de presupuestos existentes conservan su instantánea; no se modifica su SAP histórico.

Aplicados en SQL local: migración 021, los tres SP de mantenimiento de Cliente y DatosIniciales de presupuestos. Los 83 procedimientos instalados coinciden con las fuentes. Paquetes 1, 3 y 5 y sus huellas actualizados; acumulado 151 fuentes, 47 tablas, 105 relaciones. Validación offline y 155 comprobaciones SQL aisladas aprobadas, incluidas alta/edición SAP y sugerencia en nuevos presupuestos; pruebas MVC y navegador cubren campo editable y enlace oculto.

## Tiendas y contactos en menú y cliente visible — 05/11/2026

Menú Maestros incorpora Tiendas (TIENDAS.CONSULTAR) y Contactos (CLIENTES.CONSULTAR), independientes de Clientes. Conserva los accesos relacionados desde la ficha para trabajar con un cliente filtrado.

spu_MantTienda_Listar y spu_MantContactoCliente_Listar incorporan JOIN explícito a Cliente y agregan ClienteCodigo/ClienteNombre al final del primer resultado. ClienteNombre usa razón social para jurídica o nombres y apellidos para natural. Búsqueda aplica también al código/nombre del cliente antes del conteo y paginación SQL de veinte. Se mantienen permisos, filtro opcional por cliente, auditoría y orden estable por código/Id. No hay cambios de tablas.

MaintenanceRow y SqlMaintenanceRepository transportan los dos campos al listado MVC, con columnas separadas «Código de cliente» y «Cliente: nombres / razón social». Las ayudas conservan su contrato. Aplicados ambos procedimientos en SQL local; paquete y huellas de etapa 3 regenerados. Verificación: compilación, pruebas HTTP y 158 comprobaciones SQL aisladas, incluidas búsquedas por cliente jurídico y natural.

## Lista de enlaces de documentos por cliente — 05/11/2026

ClienteEnlaceDocumento almacena Id, ClienteId, Descripcion (200), Url (2048), Orden, fecha/usuario de registro y actualización. FK a Cliente y a usuarios; índice por cliente/orden/Id. Alta de tabla independiente en Tablas/ClienteEnlaceDocumento.sql, registrada antes de los SP de mantenimiento en etapa 3. Script/022_EnlacesDocumentosCliente.sql migra el antiguo UrlCargaDocumentos con descripción «Documentos del cliente»; conserva el valor anterior y marca la migración mediante propiedad extendida de tabla. Repetir el script no duplica ni restaura enlaces que el usuario haya retirado.

UrlCargaDocumentos pasa a campo oculto de compatibilidad, sin control en pantalla; no es la fuente vigente de enlaces. spu_MantCliente_Obtener devuelve Enlaces como arreglo JSON ordenado y normaliza el vacío a []. Guardar acepta Enlaces con Id opcional, Descripcion y Url: valida lista, máximos, descripción/URL, duplicados y pertenencia de los IDs; altas, modificaciones y retiros se guardan en la transacción del cliente. Conserva identidad/fecha de creación de filas editadas y asigna responsable/fecha de actualización. El bloqueo y VersionFila del cliente protegen también su lista. Un JSON antiguo sin Enlaces conserva la lista; [] retira todas las filas. Las instantáneas anterior/nueva de auditoría incluyen los enlaces y responsables, permitiendo consultar los retirados.

MVC permite agregar/quitar filas con descripción y enlace, guardándolas con «Guardar registro». Valida HTTP/HTTPS absoluto sin credenciales y hasta cien enlaces. Cada URL válida tiene un icono de apertura en otra pestaña con noopener/noreferrer; los esquemas no permitidos no producen enlaces navegables. La consulta es de solo lectura. Se mantienen filas y errores ante fallos de validación/concurrencia; los controles se adaptan con desplazamiento horizontal de la lista en móvil.

Aplicados tabla, migración y SP Obtener/Guardar en SQL local. Manifiesto acumulado: 153 fuentes, 48 tablas, 108 relaciones y 83 SP. Verificación: 165 comprobaciones SQL aisladas, incluidas varios enlaces, edición/retiro, auditoría, rechazo de enlace ajeno/esquema inválido y migración repetible; pruebas HTTP y Chrome con dos enlaces, alta/retiro/reindexado de filas, guardado y atributos de apertura segura en escritorio/móvil. Paquete de etapa 3 y huellas actualizados.

## Interfaz unificada y modo oscuro — 05/11/2026

Mejoras de presentación MVC: acciones operativas como botones, listados principales al ancho disponible, cabeceras diferenciadas de pantalla y barra roja en popups/calendarios. Fechas/responsables se unifican en franja adaptable; se retira el historial general de las ediciones y de la configuración, manteniendo consultas históricas de documentos/revisiones y todos los eventos internos. No cambia el registro SQL de auditoría.

La descripción del material usa LongText en los metadatos compilados/generador y un textarea de tres líneas, de fila completa; mantiene límite de 1500 caracteres y nombre de campo. Sin cambios de tabla, procedimiento o contratos de guardado SQL. El selector claro/oscuro en menú recuerda hyt.theme en almacenamiento local del navegador y usa la preferencia del sistema cuando no hay elección. Las ventanas embebidas sincronizan el tema mediante mensajes validados del mismo origen y los estilos de los gráficos usan colores semánticos. Impresiones HTML mantienen fondo blanco. Evidencias y alcance en VERIFICACION_INTERFAZ_HYT.md.

## Ocultación global del historial de cambios — 05/11/2026

Por corrección expresa del usuario, el historial general de cambios/acciones también se retira de las vistas de consulta de todos los maestros, presupuestos y documentos de seguimiento, incluidos los maestros embebidos. La franja Fechas y responsables sigue visible. Se conserva el almacenamiento y la generación de auditoría interna; no se modifican SQL ni datos históricos. Esta decisión sustituye el alcance anterior que conservaba la sección en consulta.

## Tarifario simple desde materiales — 05/11/2026

Firma: FRANCO LARA | Create date: 05/11/2026 | Simplificación del registro de precios del material, estado y trazabilidad propia.

Tarifario incorpora EsTarifarioMaterial BIT NOT NULL DEFAULT 0 y un índice único filtrado para permitir una sola cabecera interna gestionada por el sistema. Script/023_TarifarioMaterialDirecto.sql agrega la estructura de forma idempotente; la definición inicial y los manifiestos también están actualizados. Los registros convencionales conservan el valor 0 y sus datos.

spu_MaterialTarifa_Guardar crea automáticamente esa cabecera y una versión interna VIGENTE cuando son necesarias, sin exponerlas al usuario. Cada precio sigue siendo un registro independiente de TarifarioPartida, con su identificador, VersionFila, usuario/fecha de creación y usuario/fecha de actualización. El procedimiento valida permisos, material, IGV, PEN obligatorio, USD opcional, importes no negativos y máximo dos decimales. Usa bloqueo para la inicialización y control optimista de concurrencia para editar. Conserva las instantáneas de auditoría de alta, actualización y cambio de estado. Copia SAP y zonas del material internamente al crear.

Las tarifas del contenedor interno permiten actualización directa. Las convencionales solo permiten editar sus borradores; las publicadas conservan su protección. Listados y ayudas convencionales ocultan el contenedor interno. spu_MaterialTarifas_Listar devuelve Editable y las fechas/responsables de cada precio con paginación SQL de veinte filas. spu_MantTarifarioPartida_Obtener devuelve EsTarifarioMaterial para autorizar el formulario correspondiente.

MVC muestra Agregar tarifario con Tipo de IGV, Precio en soles, Precio en dólares opcional y Estado. No solicita versión, posición, zonas ni publicación. El material lo determina la ruta y se verifica también en el servidor. Consulta y edición muestran Fechas y responsables del precio. Los accesos embebidos anteriores redirigen a esta ventana simple. Guardar refresca el listado sin perder cambios pendientes en el formulario del material.

Los precios activos nuevos se ofrecen en presupuestos sin publicación manual, mediante las consultas vigentes. Desactivar deja de ofrecerlos. Los presupuestos existentes conservan sus instantáneas de precios; editar un tarifario no recalcula documentos ya guardados.

Verificación: 173 comprobaciones SQL en base temporal eliminada, pruebas HTTP y recorrido Chrome de alta/edición con cuatro controles; 255 comprobaciones de interfaz. Aplicado en SistemaWebHYT local y cotejados sus 84 SP con fuentes. Manifiesto: 155 scripts, 48 tablas y 108 relaciones. Documentación de uso en CODIGOS_Y_TARIFAS_DESDE_MATERIALES.md.

## Tarifarios unificados con materiales — 05/11/2026

El menú Tarifarios muestra los registros de TarifarioPartida creados desde la ficha del material y desde el propio listado. El listado convencional de cabeceras se sustituye en la navegación por /Tarifarios; /Maestros/tarifarios redirige conservando búsqueda/página. No duplica ni migra los precios existentes.

spu_MaterialTarifas_Listar admite MaterialId=0 para consultar todos los materiales y Busqueda opcional por código interno, descripción o SAP. Devuelve MaterialId, MaterialCodigo y MaterialDescripcion junto con precios, IGV, estado y responsables propios. El mismo filtro se aplica al total y al resultado paginado, con veinte filas y orden estable por fecha/identidad. Mantiene los permisos de ambos módulos y las consultas por material continúan aisladas.

Nuevo registro abre un formulario con ayuda de material y los mismos campos de la ficha: IGV, PEN, USD y estado. Ambos accesos llaman spu_MaterialTarifa_Guardar; al menos uno de PEN/USD es obligatorio. Se permite solo dólares sin inventar el precio en soles, o ambos precios manuales. La precisión sigue siendo de dos decimales. Editar/ver conserva pertenencia, concurrencia, fechas y permisos. No cambia la estructura de las tablas.

Aplicados los dos SP en SistemaWebHYT local y cotejados los 84 módulos con las fuentes. Verificación: 177 comprobaciones SQL aisladas aprobadas, pruebas HTTP y recorrido Chrome de listado unificado, ayuda de material, alta desde menú y diseño móvil sin desborde. Paquete de etapa 4 y huellas actualizados.

## Ubigeo, direcciones y datos simplificados de presupuesto — 05/11/2026

Se incorporan UbigeoDepartamentos (CodigoDepartamento CHAR(2)), UbigeoProvincias (CodigoProvincia CHAR(4), CodigoDepartamento) y UbigeoDistritos (CodigoUbigeo CHAR(6), CodigoDepartamento, CodigoProvincia), con Nombre, Activo, claves primarias y relaciones a sus padres. El catálogo se copia de FralseCont local: 25 departamentos, 196 provincias y 1889 distritos. Script/024_UbigeosCatalogo.sql contiene la carga portable e idempotente; no hay dependencia entre bases en ejecución ni escritura en FralseCont/Dbsisadm. tools/export_ubigeos.py permite preparar de nuevo la carga desde el origen local.

Script/025_DireccionesYConsideraciones.sql agrega Cliente.CodigoUbigeo y Tienda.CodigoUbigeo CHAR(6) nullable, relacionados con UbigeoDistritos, conservando registros anteriores sin asignarles distritos supuestos. Agrega ConfiguracionHYT.ConsideracionesPresupuesto NVARCHAR(MAX) y amplía PresupuestoRevision.ClienteDireccion/TiendaDireccion a NVARCHAR(600) para sus instantáneas completas.

spu_MantCatalogos devuelve también departamentos/provincias/distritos activos. Los formularios de cliente y tienda usan combos dependientes, limpian selecciones inferiores al cambiar su padre y restauran el distrito al editar/consultar. Si se registra dirección, MVC exige completar el ubigeo; las filas históricas se conservan sin inventar códigos. Los SP de alta/edición validan distritos existentes/activos y conservan CodigoUbigeo en sus instantáneas de auditoría. Los SP Obtener lo devuelven para edición. Los listados muestran Dirección / Departamento / Provincia / Distrito en la columna de dirección, con veinte filas por página desde SQL.

spu_Presupuesto_DatosIniciales devuelve automáticamente direcciones completas de cliente y tienda. La pantalla las muestra al elegir cada registro y comprueba la pertenencia de la tienda al cliente. spu_Presupuesto_Guardar conserva la dirección con ubigeo en la revisión; revisiones anteriores mantienen su instantánea. Zona informativa y Especialidad se eligen con ayudas de sus maestros, con búsqueda/paginación existente. Se conserva el texto elegido por revisión para las impresiones HTML. Trimestre ofrece 1 TRIMESTRE, 2 TRIMESTRE, 3 TRIMESTRE y 4 TRIMESTRE; valores históricos se conservan al abrir revisiones antiguas.

Encargado, Firmante, Coordinador y Cargo del firmante dejan de ser campos visibles en presupuesto. Sus valores históricos no se eliminan. Consideraciones pasa al final del formulario, antes de Guardar. En Configuración HYT se edita el texto predeterminado (hasta 20000 caracteres); los SP Obtener/Guardar incluyen el campo, concurrencia y auditoría, usando OPENJSON para no truncar textos de más de 4000 caracteres. Un presupuesto nuevo toma ese texto y permite modificarlo. Editar uno existente o elegir cliente/tienda/otras ayudas conserva las consideraciones del documento; cambiar la configuración no reescribe presupuestos guardados.

Fuentes, manifiestos de etapas 3/5 y paquetes actualizados. Aplicados en SQL local los catálogos, migraciones y SP; los 84 procedimientos coinciden con sus fuentes. Validación: 185 comprobaciones SQL aisladas, pruebas HTTP, Chrome de combos/direcciones/ayudas/trimestres/consideraciones y 255 comprobaciones de interfaz en claro/oscuro y escritorio/tablet/móvil. SQL offline: 160 fuentes, 51 tablas; sin datos de prueba en SistemaWebHYT de trabajo.

## Presupuesto: detalle por material con precios netos editables — 05/11/2026

El formulario de presupuesto ocupa todo el ancho disponible. Se retiran Aplicar datos de tienda y proveedor del cliente y Agregar desde tarifario. Partidas pasa a Detalle y el único botón de alta es Agregar ítem, rojo. Cada fila permite buscar material, buscar exclusivamente tarifas del material seleccionado, ingresar necesidad y cantidad, elegir tipo de servicio mediante ayuda, cargar automáticamente la abreviatura de unidad del material, editar precio sin IGV y consultar subtotal calculado. Las cantidades de filas nuevas comienzan vacías, obligando a ingresarlas. Las ayudas conservan búsqueda y veinte filas por página desde SQL. El detalle tiene desplazamiento horizontal interno en pantallas pequeñas.

spu_MantServicioMaterial_Obtener devuelve UnidadCodigo mediante el JOIN existente; Predeterminados de material expone unidad, código interno, SAP y descripción. spu_Presupuesto_AyudaTarifario acepta filtros opcionales MaterialId y PartidaId, además de los filtros de cliente, fecha, moneda y búsqueda existentes. Ambos se aplican antes de paginar y contar. Esto impide que la ayuda de una fila mezcle tarifas de otros materiales y permite recuperar el mismo registro al cambiar PEN/USD.

La tarifa se obtiene en la moneda del presupuesto; si su precio incluye IGV, se divide entre 1 más la tasa aplicable y se redondea a dos decimales antes de cargarlo. El precio resultante siempre es editable. No hay conversión cambiaria; si falta el precio en la moneda seleccionada, queda vacío para ingreso manual. Al cambiar moneda se vuelve a consultar la tarifa elegida y se mantiene la confirmación de revalorización para documentos existentes.

Los ítems de este formulario se guardan como precio manual neto, con referencia al material y a la tarifa elegida cuando existe; el motivo interno identifica el precio del detalle. Se conserva así la tarifa de referencia sin imponer nuevamente su precio sobre el importe editado. DetalleNeto en el contrato de guardado exige material y tratamiento SIN_IGV en MVC y SQL. La referencia tarifaria de un precio manual debe corresponder también al material, además del cliente. Los contratos antiguos siguen disponibles para compatibilidad. Las revisiones históricas no se reescriben: al abrir una para consulta/edición se presentan los precios netos; guardar una edición autorizada genera la auditoría existente.

Subtotal de ítem = cantidad por precio neto, con dos decimales. El resumen suma bases, calcula IGV y total con la política existente. Se conservan necesidad, tipo de servicio y valores de impresión en las instantáneas de cada revisión. No se agregan tablas ni columnas en esta modificación. Los tres SP afectados se aplicaron en SistemaWebHYT local; 84 módulos cotejados con las fuentes. Paquetes de etapas 3 y 5 actualizados.

Verificación: 190 comprobaciones SQL en base aislada, pruebas HTTP y recorrido Chrome de material/unidad, tarifa CON IGV convertida a precio neto, edición de precio, cantidad explícita, servicio/necesidad, subtotal e IGV separados y cambio a moneda sin precio. Compilación sin errores. Evidencias ficticias: entregables/preview/presupuesto-detalle-neto-escritorio.png y presupuesto-detalle-neto-movil.png.

## Numeración mensual y detalle compacto — 06/11/2026

Presupuesto agrega NumeroPresupuesto NVARCHAR(40) NOT NULL con índice único UX_Presupuesto_NumeroPresupuesto. PresupuestoId y NumeroInterno conservan sus valores y generación anteriores. La nueva tabla interna PresupuestoCorrelativoMensual usa Periodo CHAR(6) como clave y Ultimo BIGINT como contador no negativo. No se expone mantenimiento de este contador.

Script/026_NumeroPresupuestoMensual.sql asigna números a registros existentes por año/mes de su FechaRegistro convertida de UTC a Lima, ordenados por fecha e identificador, y prepara los contadores. La carga se realiza con bloqueo y transacción; repetir el script conserva números ya asignados y nunca reduce contadores. Se mantiene identidad, enlaces de documentos y auditoría existente. Las impresiones HTML conservadas anteriormente no se reescriben.

spu_Presupuesto_Guardar reserva el siguiente correlativo del mes con UPDLOCK/HOLDLOCK dentro de la misma transacción que registra el presupuesto. Usa un único instante UTC para determinar el mes de Lima y guardar FechaRegistro. Formato visible: yyyyMM seguido de correlativo con al menos seis posiciones; por ejemplo 202611000001. El contador se reinicia por mes/año y el número no cambia al editar ni al crear revisiones. Antes de guardar, la tarjeta indica Se asigna al guardar; no se reserva un número solo por abrir la ventana.

Los SP de listado/consulta de presupuesto, seguimiento OC/guía/confirmación/ejecución, dashboard y reportes devuelven NumeroPresupuesto usando los nombres de columna de sus contratos existentes. Las impresiones nuevas de ambos formatos toman ese número desde la consulta del presupuesto; se conserva compatibilidad con DTO y contratos sin cambiar el identificador de relaciones. Las búsquedas de esas listas utilizan el número visible.

El detalle usa botones pequeños de lupa para material, tarifa y tipo de servicio; quitar ítem usa un tacho. Mantienen títulos, nombres accesibles y foco de teclado. El campo necesidad se reduce aproximadamente un 70 %. La tarifa del material abre sin seleccionar cliente, ofreciendo solo tarifas genéricas; al elegir cliente también se muestran las autorizadas para este. Los filtros de material, fecha y moneda y la paginación SQL de veinte filas continúan vigentes. Las tarifas privadas de otros clientes no se ofrecen en la consulta sin cliente.

Ver y Editar tienen ancho uniforme en todas las grillas compartidas y listas personalizadas. Los códigos internos conservan su ubicación debajo del retorno y se presentan en una tarjeta con borde superior rojo. Presupuesto utiliza la misma tarjeta para el número visible. Estos cambios respetan los temas claro/oscuro y el desplazamiento interno del detalle en móvil.

Fuentes y manifiesto de etapa 5 incluyen tabla y migración antes de los SP dependientes. Paquetes de etapas 5/6/8 y huellas actualizados. Verificación: 195 comprobaciones SQL aisladas, incluidos correlativos concurrentes, correspondencia con mes de registro en Lima, migración repetible y exclusión de tarifas privadas sin cliente. Pruebas HTTP y Chrome aprobaron apertura sin cliente, iconos compactos, anchos uniformes y tarjetas. Evidencias ficticias en entregables/preview/presupuesto-detalle-compacto-escritorio.png y presupuesto-detalle-compacto-movil.png. Se aplicó en SQL local y se cotejaron los 84 procedimientos con sus fuentes.

## Nombre del tipo de servicio y columnas compactas — 06/11/2026

La selección de Tipo de servicio en el detalle carga y conserva la descripción del maestro, en lugar de su código. La ayuda de tarifas también devuelve la descripción de servicio y el guardado de una tarifa tradicional usa ese nombre como predeterminado cuando corresponde. PresupuestoPartida.TipoServicio pasa de NVARCHAR(60) a NVARCHAR(150), acorde con el maestro. Script/027_NombreTipoServicioPresupuesto.sql amplía la columna sin modificar valores históricos; se actualizan la definición inicial, las tablas temporales/lectura JSON y el límite de validación de spu_Presupuesto_Guardar, además de la validación MVC. Las nuevas instantáneas e impresiones conservan el nombre completo. No se reinterpretan ni reescriben los códigos de presupuestos históricos.

Código Material SAP se muestra después del nombre del material, manteniendo su campo de guardado y carga automática desde material/tarifa. Se elimina su control oculto duplicado. Unidad se renombra Unidad de medida. Servicio se compacta aproximadamente 60 % y unidad/cantidad/precio reducen su ancho; clases semánticas de columna mantienen estos tamaños al insertar SAP. Los valores completos siguen disponibles al enfocar o consultar su título. Se conserva el desplazamiento horizontal interno en móvil.

Al cambiar moneda, una respuesta pendiente de tarifa no reemplaza un precio que el usuario ya ingresó manualmente mientras esperaba. Validación: 196 comprobaciones SQL aisladas, incluidas conservación de nombre de cien caracteres y regresiones existentes; pruebas HTTP y Chrome de selección por nombre, posición/carga de SAP, anchos y cálculo. Fuentes, paquetes de etapas 1/5 y huellas actualizados. Script 027 y SP de guardado/ayuda aplicados en SQL local.

06/11/2026 — Ajuste del flujo del detalle: se retira la casilla de confirmación de revalorización y el aviso repetitivo de cantidades durante la captura. Mientras falta una cantidad válida la previsualización queda pendiente; se mantienen la obligatoriedad y validación al guardar. El controlador autoriza automáticamente la revalorización del formulario DetalleNeto al cambiar moneda. Los precios del detalle siguen en SIN IGV; tarifas CON IGV se convierten dividiendo entre 1 más la tasa aplicable, a dos decimales. Sin cambios SQL ni de estructura.

## Configuración HYT con ubigeo y títulos de maestros — 06/11/2026

Los listados de clientes y tiendas cambian el encabezado de su dirección completa a Dirección; el contenido sigue incluyendo dirección y ubigeo. En tiendas, Cliente: nombres / razón social pasa a Cliente. No se modifican las búsquedas, datos ni columnas SQL de esos listados.

ConfiguracionHYT agrega CodigoUbigeo CHAR(6) NULL, relacionado por FK_ConfiguracionHYT_Ubigeo con UbigeoDistritos. Un único código de distrito determina departamento y provincia mediante el catálogo; no se duplican nombres ni selecciones contradictorias en la tabla. La definición inicial y Script/028_UbigeoConfiguracionHYT.sql cubren bases nuevas y existentes. La migración puede repetirse y conserva los datos existentes, sin asignar ubicaciones supuestas.

La configuración muestra, después de Dirección, combos dependientes de Departamento, Provincia y Distrito con el catálogo ya cargado en HYT. Reutiliza los catálogos y comportamiento de clientes/tiendas: cambiar departamento/provincia limpia las selecciones inferiores; al editar se restaura el distrito y sus padres. El formulario envía Fields[CodigoUbigeo], conserva la selección ante errores y respeta el permiso de edición.

spu_ConfiguracionHYT_Obtener devuelve CodigoUbigeo. spu_ConfiguracionHYT_Guardar valida longitud y pertenencia a un distrito activo, persiste el código con la dirección y lo incluye en las instantáneas anterior/nueva de auditoría. Se mantiene el control de concurrencia de la configuración. No cambia el contenido de presupuestos ni impresiones conservadas.

Fuentes, generador de esquema y manifiestos/paquetes de etapas 1/3/5 actualizados. Aplicados Script 028 y ambos SP de configuración en SQL local y cotejados los 84 módulos con sus fuentes. Pruebas: 199 comprobaciones SQL aisladas, incluidas persistencia/auditoría de ubigeo, rechazo de distrito inexistente y migración repetible; HTTP y Chrome aprobaron títulos, combos, restauración de selección, envío de distrito y presentación móvil. Manifiesto: 164 fuentes y 52 tablas.

## Estados y listado de presupuestos — 06/11/2026

Nombres visibles: Registrado, Enviado, Observado, Aceptada, Rechazada y Cancelado. Se conservan códigos SQL BORRADOR/APROBADO/RECHAZADO para compatibilidad con OC, guías, reportes e impresiones históricas. Sin cambios de estructura ni reescritura de auditoría previa. spu_Presupuesto_Guardar permite editar la última revisión OBSERVADO sin OC, con rowversion, recálculo y auditoría del estado real; no genera otra revisión ni permite modificar una revisión aceptada/enviada. spu_Presupuesto_CambiarEstado permite reenviar OBSERVADO y bloquea VENCIDO como destino.

Desde el listado, Cambiar estado abre el popup HYT con las transiciones permitidas y motivo condicional. CSRF, autorización y control de concurrencia permanecen activos. El listado retira Registro, Registrado por, Actualización y Actualizado por; la consulta conserva Fechas y responsables. La fecha de aceptación permanece en el listado.

Los nuevos nombres también se muestran en gráficos, detalle accesible, reporte de presupuestos ganados y nuevas impresiones HTML (plantilla HYT-2026.2). Las impresiones previamente conservadas mantienen su contenido original.

## Firma de presupuesto y formato nuevo — 06/11/2026

`ConfiguracionHYT.FirmaPresupuestoJson NVARCHAR(2000) NULL` conserva únicamente NombreOriginal, ClaveAlmacenamiento, TipoContenido, TamanoBytes y SHA256. Imagen privada JPG/PNG en Cloudflare R2; máximo 2 MB y 4096 píxeles por lado. La migración idempotente es Script/029_FirmaPresupuesto.sql y la tabla base incluye el campo. `spu_ConfiguracionHYT_FirmaGuardar` exige CONFIGURACION.EDITAR, versión original, metadatos válidos y registra ACTUALIZAR_FIRMA con valores anterior/nuevo; guardar los datos corporativos no borra la firma. Obtener incorpora los metadatos y permite lectura de datos corporativos a usuarios con PRESUPUESTOS.CONSULTAR y REPORTES.CONSULTAR para generación de documentos, además de CONFIGURACION.CONSULTAR. La pantalla de configuración mantiene su autorización propia.

MVC: Configuración HYT → Subir firma para presupuesto, POST multipart con CSRF y rowversion. Vista previa mediante endpoint autorizado, sin URL pública de bucket. Generar impresión recupera la imagen desde R2, verifica tamaño/SHA256 y la incorpora como data URI al HTML conservado. Si falla recuperar una firma configurada, informa el error. Las impresiones previamente conservadas no se reescriben ni dependen de enlaces temporales. Plantilla HYT-2026.3.

El formato NUEVO se reconstruye contra hytpresupuesto.pdf: cabecera serif, cuadros de cliente/número/fecha, trimestre/zona/vigencia, proveedor, tabla de once columnas roja, subtotal amarillo, consideraciones sin texto automático duplicado y firma al pie. No imprime estado, revisión, fecha de generación, tipo de atención ni otros datos internos. Mantiene cálculo real y moneda; no reproduce #N/A ni valores ficticios de la planilla. El formato anterior conserva su estructura y admite la firma.

## Posición, centro FICO y subtítulo de cotización — 06/11/2026

`ConfiguracionHYT.SubtituloCotizacion NVARCHAR(2000) NULL` contiene el subtítulo completo de la cotización, incluidos teléfonos y los dos correos. Campo LongText en Configuración HYT. Valor inicial: texto proporcionado por el usuario, mediante DEFAULT y migración idempotente Script/030_SubtituloCotizacion.sql; no reemplaza ediciones al reaplicar. `spu_ConfiguracionHYT_Guardar/Obtener` valida, persiste, consulta y audita el campo. La omisión por un cliente anterior conserva el valor existente; un texto vacío explícito lo deja vacío.

El generador toma el subtítulo vigente al crear cada impresión y lo centra a todo el ancho, en ambos formatos. No se concatena otra vez teléfono/correo ni se conserva texto de servicios fijo en las vistas. El HTML ya emitido es inmutable. Plantilla HYT-2026.4.

`spu_Presupuesto_Guardar` asigna Posicion y CodigoInterno desde ServicioMaterial.CodigoInterno para cada material elegido; CentroSAPFICO se asigna desde Tienda.CentroSAPFICO. Se ignoran esos valores cuando vienen manipulados del navegador. Para impresiones nuevas de revisiones ya registradas, la columna POSICION usa CodigoInterno conservado en el ítem. No se reescriben partidas, códigos, maestros ni impresiones históricas en la migración.

## Respaldo de Centro SAP FICO en presupuestos anteriores — 06/11/2026

`spu_Presupuesto_Obtener` devuelve CentroSAPFICO de la partida cuando tiene valor; si es NULL, vacío o espacios, usa Tienda.CentroSAPFICO de la tienda vinculada al presupuesto. La consulta no actualiza partidas ni impresiones conservadas. Esto permite generar impresiones nuevas correctas de presupuestos guardados antes de configurar el centro en la tienda, sin editar ni volver a guardar el presupuesto. El guardado nuevo continúa tomando el centro desde la tienda en SQL.

## Formato anterior e impresión desde encabezado — 06/11/2026

Plantilla HYT-2026.8: ANTERIOR se adapta al Excel H Y T - PPTO 100(ANT).xlsx, manteniendo siete columnas (ítem, descripción, unidad, cantidad, precio unitario, precio parcial y total del grupo), cabecera roja y costo final sin IGV destacado en rosado. Aplica Arial, subtítulo completo centrado con línea roja, número íntegro de presupuesto/fecha, cliente/dirección, tienda, código de proveedor, consideraciones editables y firma R2. No imprime descripción interna del trabajo, coordinador/encargado, revisión, estado, generación ni notas automáticas contradictorias; evita repetir despedidas ya incluidas en consideraciones. Importes del detalle anterior siempre netos; revisiones antiguas CON_IGV convierten el precio unitario con su tasa histórica y usan BaseImponible para parciales y grupos. Mantiene IGV y total calculados en revisión. No cambia el registro original ni impresiones ya emitidas.

La sección Impresión del presupuesto se mueve a una subtarjeta dentro del encabezado Presupuesto, con Imprimir formato nuevo e Imprimir formato anterior. Disponible al consultar o editar un presupuesto guardado, con permiso REPORTES.CONSULTAR; ausente en alta sin guardar y revisiones nuevas sin guardar. Ambos botones usan POST con CSRF y versión original, y abren una pestaña nueva. Se oculta el enlace Consultar impresiones conservadas sin eliminar su almacenamiento ni controles de acceso.

06/11/2026 — ajuste de ANTERIOR: se retiran las dos filas de resumen IGV y TOTAL (cuatro celdas visibles), conservando COSTO FINAL SIN IGV. La cabecera muestra TiendaNombre sin prefijo TiendaCodigo. Los impuestos e importes continúan calculados y guardados; el formato NUEVO no cambia. Plantilla HYT-2026.9.

06/11/2026 — contacto en formato ANTERIOR: `spu_Presupuesto_Obtener` devuelve ContactosTienda, concatenando nombres de contactos activos del mismo ClienteId y TiendaId, por ContactoId y separados con / si hay varios. No incluye contactos generales del cliente, inactivos ni de otras tiendas. La cabecera lo muestra debajo de Tienda como Coordinador/Encargado:. No utiliza los campos internos Coordinador/Encargado del presupuesto. Sin cambios de estructura ni reescritura de HTML emitido. Plantilla HYT-2026.10.

06/11/2026 — captura OC: muestra número de presupuesto en tarjeta derecha con borde superior rojo y Estado del presupuesto informativo. Se retira el selector de revisión; el servidor resuelve la revisión asociada (vigente en altas, conservada en ediciones) y mantiene las comprobaciones de versión y permisos. Sin cambios SQL.

## Aceptación automática por OC o Guía — 06/11/2026

`spu_SeguimientoOC_Guardar` y `spu_SeguimientoGUIA_Guardar` aceptan automáticamente el presupuesto al crear el documento, para invitación y emergencia, desde Registrado, Enviado u Observado. Conservan Aceptada si ya lo estaba. OC exige revisión vinculada válida e importes completos; Guía ya no exige OC previa. Rechazada, Vencido heredado y Cancelado no se aceptan automáticamente. Las ediciones de documentos no alteran el estado del presupuesto.

Se actualizan Estado (APROBADO), FechaAprobacion (fecha actual de Perú), FechaActualizacion y UsuarioActualizacion de la revisión; el encabezado conserva su actualización de responsable/fecha. Auditoría CAMBIAR_ESTADO registra estado anterior/nuevo, usuario, correlación, motivo automático y documento de origen. Todo ocurre en la misma transacción que documento, adjuntos y operación idempotente: cualquier fallo revierte la creación y la aceptación; reintentar no duplica auditoría. No altera el estado de conformidad de la Guía ni la confirmación del cliente de la OC. No hay cambios de tablas ni migración de presupuestos existentes.

## Consulta compacta OC y adjuntos — 06/11/2026

`spu_SeguimientoOC_Listar` devuelve GuiaId, NumeroGuia y AdjuntosGuia (solo archivos vigentes), mediante la relación única presupuesto/guía. Solo presenta la guía si el usuario tiene GUIAS.CONSULTAR; conserva la paginación SQL de 20, filtros, conteo y permisos originales. Sin cambios de tablas.

Listado OC: se retiran Fecha comprobante, Tipo, Serie y Número; Facturado Sí abre consulta informativa por doble clic o Enter. Usa GET /OCCliente/{id}/Facturado/Consulta con permisos FACTURADO.CONSULTAR y OC_CLIENTE.CONSULTAR y consulta vigente al servidor. No permite modificar desde esa consulta. Columnas Adjuntos, Adjuntos de confirmación y nueva Guía / Adjuntos de guía abren GET /Seguimiento/{tipo}/{id}/Adjuntos, validando el permiso propio del documento. La ventana muestra nombre, fecha de carga, responsable y enlaces de visualización/descarga; no permite retirar archivos. La descarga conserva validación de permiso y almacenamiento privado R2. No genera URLs públicas.

## Documentos registrados desde presupuesto — 06/11/2026

`spu_Presupuesto_Listar` incorpora EjecucionId mediante LEFT JOIN por PresupuestoId (relación única UX_Ejecucion_Presupuesto), conservando las columnas y agregando el ID al final para el mapeo ADO.NET. Sin cambios de tablas, filtros ni paginación de 20. BudgetListItem añade ExecutionId opcional.

Guía, OC del cliente y Ejecución aparecen con borde rojo y marca de registro si su ID existe, independientemente de que tengan adjuntos. El clic abre la ficha de consulta, donde se muestran detalle y adjuntos privados disponibles para visualizar/descargar. Sin registro, usuarios con edición abren el alta en pop-up; usuarios sin consulta no reciben la acción. Se conservan todos los permisos y controles de descarga.

## Carpetas privadas de archivos en Cloudflare R2 — 06/11/2026

Las nuevas cargas usan OCCliente/{NumeroOCCliente seguro}/{operacion}.pdf|png|jpg, Guia/{NumeroActa seguro}/{operacion}.pdf, ConfirmacionCliente/{NumeroConfirmacion seguro}/{operacion}.pdf|png|jpg y ConfigHYT/Firma/{operacion}.png|jpg. El generador también admite Presupuestos/{numero}/{operacion} para cargas de presupuesto futuras; la impresión HTML actual sigue su almacenamiento existente. No se agrega un formulario nuevo de adjuntos al presupuesto.

Los segmentos conservan ASCII A-Z, a-z, 0-9, guion y guion bajo; otros caracteres se reemplazan por _. La función dbo.fn_ArchivoSegmento (registrada en manifiesto general y etapa 6) comparte esa regla con DocumentUploadRules.Folder. Los SP spu_SeguimientoOC_Guardar, spu_SeguimientoGUIA_Guardar y spu_SeguimientoCONFIRMACION_Guardar validan exactamente la carpeta del número de documento y el nombre GUID de operación/extensión. spu_ConfiguracionHYT_FirmaGuardar admite el prefijo fijo ConfigHYT/Firma. Las validaciones de tipo, tamaño, integridad, idempotencia, concurrencia, permisos y auditoría continúan vigentes. Sin cambios de tablas.

ClaveAlmacenamiento conserva la ruta completa; R2 y almacenamiento privado local soportan subcarpetas. Referencias sin carpeta de archivos históricos siguen legibles y válidas; no se mueven ni renombran objetos existentes. Si cambia el número del documento, las cargas nuevas usan la carpeta del nuevo número, y las anteriores se consultan por su clave original. Los nombres originales permanecen en metadatos, y el nombre físico GUID evita sobreescrituras. Todos los accesos siguen privados y autenticados; no se generan URLs públicas ni se cambia el bucket.

## Guía sin flujo de estados — 06/11/2026

Guía es un documento con fecha, número alfanumérico, observación y PDF. Ver/Editar reutilizan _Form con DocumentFormModel.ReadOnly: consulta bloqueada sin guardar/cargar/retirar adjuntos ni motivo de edición; edición con los mismos datos, motivo, fechas y responsables. Se elimina filtro/columna Estado de su listado y toda transición visible. El acceso a Guía desde OC consulta el mismo formulario bloqueado. Se conservan adjuntos privados, permisos y auditoría.

La ruta web Estado y spu_SeguimientoGUIA_CambiarEstado rechazan transiciones de Guía. spu_SeguimientoGUIA_Guardar ya no bloquea correcciones por estados previos; mantiene concurrencia, motivo de edición, permisos, expediente marcado facturado, validación PDF y aceptación automática del presupuesto al crear Guía. Las columnas históricas Estado/FechaAceptacion/RepresentanteCliente se conservan sin migrar datos; su estado deja de determinar el flujo operativo.

spu_Facturado_Guardar y tr_ReferenciaComprobante_Requisitos (Script/007_ProteccionAuditoria.sql) exigen Guía registrada, trabajo terminado y confirmación vigente, sin aceptar la Guía por separado. spu_SeguimientoOC_CambiarEstado exige Guía existente para cerrar y ya no pide anularla para cancelar; mantiene confirmación y ejecución. spu_SeguimientoEJECUCION_CambiarEstado y spu_Presupuesto_CambiarEstado eliminan los bloqueos basados en estados de Guía. spu_Dashboard_Obtener y spu_Reporte_Listar usan existencia de GuiaId para presencia, pendientes y lista para marcar. Sin cambios de tablas ni actualización masiva de registros históricos.

## KPI y filtros de listados — 06/11/2026

spu_Presupuesto_Listar agrega @TipoAtencion (INVITACION/EMERGENCIA), @PendienteGuia y @PendienteOC. Los checks significan ausencia de documento registrado. spu_SeguimientoOC_Listar agrega @PendienteConfirmacion (sin confirmación vigente) y @PendienteFacturar (sin referencia de comprobante vigente); este último requiere FACTURADO.CONSULTAR también en SQL. Los filtros combinados se aplican mediante AND. No se agregan exclusiones implícitas por estado.

Ambos procedimientos retornan filas paginadas de 20, total y un tercer resultado Estado/Cantidad (COUNT_BIG). Los KPI consideran todo el conjunto de búsqueda, fechas, tipo y pendientes, antes del filtro Estado; este filtra la grilla y su total. ADO.NET consume los tres resultados. No hay cambios de tablas. Los estados sin resultados muestran cero: seis tarjetas en presupuesto y cuatro en OC. Las tarjetas siguen el borde superior rojo, conteo centrado y tema claro/oscuro.

Se retiran los controles Cliente/Tienda en Presupuestos, OC y Guía; Buscar continúa incluyendo cliente/tienda y se conserva compatibilidad con enlaces anteriores del dashboard. El contador de adjuntos de Guía abre la consulta bloqueada del documento y sus opciones privadas de visualización/descarga. Script incremental: entregables/ACTUALIZAR_KPI_FILTROS_LISTADOS.sql; fuentes incorporadas en paquetes etapas 5/6.

## Confirmación del cliente y trabajo completado de OC — 06/11/2026

spu_SeguimientoCONFIRMACION_Guardar cambia una OC REGISTRADA a TRABAJO_COMPLETADO al guardar una confirmación vigente, tanto nueva como editada. No requiere ejecución terminada ni permiso OC_CLIENTE.EDITAR adicional al permiso CONFIRMACIONES.EDITAR, porque es una consecuencia de la confirmación autorizada. Conserva estados TRABAJO_COMPLETADO/CERRADA_OPERATIVAMENTE; OC canceladas y expedientes facturados mantienen sus bloqueos. Actualiza usuario/fecha y registra CAMBIAR_ESTADO_AUTOMATICO en OrdenCompraCliente con estado previo/nuevo, fecha de confirmación, motivo y datos de la confirmación. Todo se ejecuta en la misma transacción: si falla la auditoría se revierte confirmación, archivos y cambio de estado; reintentos idempotentes no duplican eventos. Se mantienen versiones de concurrencia.

spu_SeguimientoOC_CambiarEstado admite Trabajo completado con confirmación vigente o ejecución terminada. El cierre operativo conserva ejecución terminada y guía registrada; cancelaciones, reaperturas y marcado facturado mantienen requisitos existentes. No cambia automáticamente la ejecución ni crea documentos de ejecución. Confirmaciones históricas no se actualizan masivamente: una OC ya confirmada puede pasar manualmente a Trabajo completado desde su listado o al guardar su confirmación. Sin cambios de tablas. Entregable ACTUALIZAR_OC_CONFIRMACION_AUTOMATICA.sql y etapa 6 sincronizados; parches anteriores incluyen definiciones vigentes.

## Facturado con OC completada — 06/11/2026

spu_Facturado_Guardar y tr_ReferenciaComprobante_Requisitos (Script/007_ProteccionAuditoria.sql) admiten trabajo acreditado por OC en TRABAJO_COMPLETADO/CERRADA_OPERATIVAMENTE o, para compatibilidad, ejecución TERMINADA. Una OC completada no necesita ejecución separada; conserva guía registrada y confirmación vigente perteneciente a esa OC como requisitos obligatorios. OC canceladas se rechazan. La confirmación, versiones, permisos, transacción, auditoría y documentos 01/03 se mantienen. No se genera ni cambia una ejecución automáticamente. La ruta alternativa basada en ejecución mantiene el bloqueo por ejecución pendiente del trigger.

spu_Dashboard_Obtener (ListasParaMarcar) y spu_Reporte_Listar (completadas/pendientes de facturar y etiqueta lista para marcar) reconocen el mismo estado de OC como acreditación del trabajo. Sin cambios de tablas ni actualización masiva de datos. Patch ACTUALIZAR_FACTURADO_OC_COMPLETADA.sql; paquetes de instalación e incrementales sincronizados.

## Reportes integrales y dashboard vigente — 06/11/2026

Normativa funcional consolidada: FLUJO_FUNCIONAL_ACTUAL_HYT.md. Esta sección reemplaza la clasificación inicial de reportes ganados/completadas/con acta/pendientes en la pantalla; las notas anteriores se conservan como evolución histórica.

spu_Reporte_Listar agrega @Estado, @TipoAtencion, @Busqueda, @PendienteGuia, @PendienteOC, @PendienteConfirmacion, @PendienteFacturar y @Exportar (defaults compatibles). Nuevos códigos PRESUPUESTO/OC_CLIENTE. Devuelve filas JSON, total y agrupación de importes por moneda. Estados permitidos dependen de fuente; joins uno-a-uno a OC/Guía/Ejecución y vínculos vigentes a confirmación/referencia, conservando una fila por documento. Columnas nuevas: PresupuestoNumero/PresupuestoRegistro/TipoAtencion/EstadoPresupuesto, OCInterno/NumeroCliente/OCRegistro/EstadoOC, Guia/TipoGuia/GuiaRegistro/GuiaFecha, EjecucionId/Registro/Inicio/Termino, Confirmacion/Registro/Fecha, TipoComprobante/Serie/NumeroComprobante/FacturadoRegistro/ComprobanteFecha/Facturado. Sin columnas de última actualización ni usuarios de auditoría en el reporte. No hay cambios de tablas.

La cohorte usa registro del presupuesto u OC según fuente, no el registro de las relaciones. @Exportar=1 devuelve todo el conjunto filtrado; caso normal conserva 20 filas. Pendientes de confirmación/facturar excluyen OC canceladas. Referencias se proyectan solo con permiso del módulo relacionado; filtros pendientes también exigen dicho permiso. Se conservan códigos/parámetros anteriores únicamente en SQL/ADO.NET para compatibilidad, fuera del selector web.

spu_Dashboard_Obtener incorpora PresupuestosSinOC, PresupuestosSinGuia y PendientesFacturar, conciliados con los filtros nuevos. Conserva estados, totales por moneda, primeros eventos y bandeja histórica. La interfaz retira filtros comerciales visibles, presenta seis estados de presupuesto/cuatro de OC y sustituye accesos a categorías viejas por reportes integrales. MVC: ReportsController filtra/valida y usa Print para HTML autónomo, _Results compartido, ReportFilter en Application y parámetros tipados ADO.NET. Exportación y consultas validan permisos por igual. Todos los archivos son UTF-8/CRLF; SDK incluye nuevas vistas/clases automáticamente.

## Dashboard: Confirmación del cliente y Facturado — 06/11/2026

Sustituye los indicadores temporales de primeros eventos del diseño anterior. Los controles, leyendas y detalle se denominan Confirmación del cliente y Facturado. spu_Dashboard_Obtener ya no calcula MIN(FechaRegistro) ni agrupa la primera incorporación histórica por OC: #Eventos contiene confirmaciones/referencias Vigente=1 asociadas a OC no canceladas, filtradas por su propia FechaRegistro y por contexto comercial. La curva agrupa esos registros por fecha de Lima. Se mantienen claves JSON Confirmaciones/Marcados por compatibilidad de contrato; no son nombres visibles. Presupuestos/Ordenes se mantienen por su fecha de registro y concilian con sus KPI del período. El resto de estados y pendientes conserva reglas vigentes.

Sin cambios de tablas. Patch ACTUALIZAR_DASHBOARD_CONFIRMACION_FACTURADO.sql y paquetes anteriores sincronizados. La documentación funcional vigente especifica que se excluyen registros retirados, no los reemplaza un primer registro histórico.

## Administración de usuarios, perfiles de acceso y auditoría · 07/11/2026

Desde **Administración → Usuarios y permisos → Nuevo usuario**, el administrador registra nombres, apellidos, correo, perfil, estado y contraseña temporal. La pantalla permite mostrarla y copiarla al guardarla. El correo se usa como inicio de sesión para las nuevas cuentas. El administrador comunica las credenciales por el medio habitual de HYT.

El primer ingreso y el ingreso posterior a un restablecimiento abren **Cambia tu contraseña**. Esa sesión puede cambiar la contraseña o cerrarse. Los adjuntos y las acciones de negocio se bloquean hasta completar el cambio. **Cambiar contraseña** está en la cabecera para cualquier perfil. El cambio voluntario requiere la contraseña actual. Las nuevas contraseñas tienen entre 12 y 128 caracteres.

| Datos | Cambios |
| --- | --- |
| AspNetUsers | Apellidos, DebeCambiarClave, FechaCambioClave, UltimoIngreso, UsuarioRegistro y UsuarioActualizacion. Índice único para el correo normalizado. Los nombres, hashes y el acceso de las cuentas existentes se conservan. |
| AspNetRoles | FechaRegistro, FechaActualizacion, UsuarioRegistro y UsuarioActualizacion. Los perfiles pueden duplicarse con sus permisos y asignarse a usuarios. |
| RolPermiso y Permiso | Acciones detalladas por módulo. Se amplían los perfiles existentes en la migración inicial y se incorpora Administración. Volver a aplicar el paquete respeta las asignaciones posteriores. |
| EventoAuditoria | Se mantiene su protección de solo lectura para la aplicación. Los nuevos eventos guardan actor, registro, motivo y cambios pertinentes. Las contraseñas y sus hashes quedan fuera del registro. |

| Acción | Procedimientos y autorización |
| --- | --- |
| Consultar usuarios y perfiles | spu_Usuarios_Listar, spu_Usuarios_Obtener, spu_Perfiles_Listar, spu_Perfiles_Obtener y spu_Seguridad_Catalogos. Búsqueda y veinte filas por página. |
| Alta y edición de usuarios | spu_Usuarios_Guardar valida correo, perfil y versión. USUARIOS.CREAR y USUARIOS.MODIFICAR son permisos independientes. |
| Activar, desactivar, restablecer, desbloquear y cerrar sesiones | spu_Usuarios_Accion valida permiso específico, motivo, versión y alcance sobre la cuenta. Rota el sello de seguridad y registra el evento. spu_RestablecerAcceso también exige cambiar la contraseña después del restablecimiento. |
| Guardar o duplicar perfiles | spu_Perfiles_Guardar valida casillas y dependencias. Un bloqueo de aplicación serializa los cambios de usuarios y perfiles. spu_Seguridad_ProtegerAdministrador comprueba que sobreviva una cuenta activa con acceso y permisos administrativos esenciales. La aplicación relee permisos en cada solicitud. |
| Cambiar contraseña propia | spu_Usuario_CambiarClave compara el sello, el hash previo y el indicador obligatorio. Actualiza el hash, revoca sesiones anteriores y registra el cambio. spu_ConsultarAcceso devuelve el indicador y spu_FinalizarAcceso registra último ingreso y bloqueo por intentos. |
| Auditoría general | spu_Auditoria_Listar filtra por fechas, actor, módulo, acción y referencia. El módulo Presupuestos reúne los eventos de cabecera y revisión. La pantalla ofrece las acciones registradas para el módulo en el período elegido. spu_Auditoria_Obtener recupera responsable, motivo y datos de un evento. La vista usa fechas de Perú y presenta Antes/Después con ocultamiento de campos sensibles. |
| Maestros | spu_Mant*_Guardar y spu_MaterialTarifa_Guardar comprueban Crear, Editar y Activar cuando se cambia la vigencia. Los contactos se administran con Clientes. Especialidades y tipos de servicio se administran con Materiales. Publicar tarifa exige editar y copiar versión exige crear. |
| Presupuestos y seguimiento | spu_Presupuesto_Guardar, spu_Presupuesto_CambiarEstado, spu_Seguimiento*_Guardar y spu_Seguimiento*_CambiarEstado validan el permiso de la operación y preservan las reglas comerciales. Facturado conserva su permiso y requisitos. |
| Impresión y adjuntos | spu_Presupuesto_RegistrarImpresion, spu_Presupuesto_ObtenerImpresion, spu_Reporte_Listar y spu_Seguimiento_ObtenerArchivo diferencian impresión y descarga de la consulta. |

El paquete incremental es **entregables/ADMINISTRACION_HYT.sql**. Se ejecuta sobre una base HYT actualizada a las etapas previas. Usar SQLCMD con detención al primer error para que el paquete pueda revertirse completo si falla. Las fuentes y huellas están en **Basededatos/Dbsishyt/manifest-administracion.json** y **entregables/SHA256_ADMINISTRACION.json**. Después de la base, iniciar la versión web que incluye la administración.

Para comprobar en una base aislada, ejecutar Hyt.DatabaseTests con --run-local-sql --use-bundles. La suite Hyt.AccessTests verifica la sesión restringida, el cambio obligatorio y voluntario, la navegación administrativa y el detalle de auditoría.
