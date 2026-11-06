# Verificación de alcance — etapas 5 y 6

Fecha: 03/11/2026. Fuentes vigentes, resultados y entregables inspeccionados antes de cerrar la implementación. La creación del bucket y colocación de claves Cloudflare fue diferida expresamente por el usuario.

| Requisito | Evidencia vigente | Resultado |
|---|---|---|
| MVC C#, ADO.NET y SP; sin EF | Solución, Program.cs, repositorios SqlBudgetRepository/SqlDocumentRepository | Implementado; compilación válida |
| Una tienda y atención Invitación/Emergencia | Presupuesto_Guardar, BudgetController, campos MVC y pruebas SQL | Implementado |
| Campos de los dos Excel y códigos alfanuméricos | ANALISIS_FORMATOS_PRESUPUESTO.md, BudgetDraft, formulario, plantillas nuevo/anterior | Conciliados y renderizados |
| PEN/USD, precio USD opcional, sin conversión automática | BudgetAmounts, SP, ayuda y pruebas SQL/Chrome | Pendiente de precio no equivale a cero |
| CON/SIN IGV, tasa por vigencia y snapshot | TasaIGV, guardado y cálculo; pruebas de ambas modalidades | Implementado |
| Dos decimales en cantidades, precios e importes | Reglas C#/SQL, restricciones, plantillas y pruebas | Implementado |
| Borradores, revisión inmutable enviada y nuevas revisiones | Presupuesto_Guardar/CambiarEstado y pruebas SQL | Implementado |
| Bloqueo de contenido con OC y regularización de emergencia | SP CambiarEstado, UI y casos SQL/MVC/Chrome | Solo enviar/aprobar; importe conservado |
| Búsqueda, último mes, fechas FralseCont y página SQL de veinte | Listados/ayudas compartidos; pruebas HTTP/Chrome | Implementado |
| Dos formatos HTML, una/varias páginas, textos largos | Siete casos Chrome + check_report_pdfs.py | Nuevo/anterior y paginación verificados |
| Impresión histórica con formato/versión/usuario y contenido íntegro | ImpresionPresupuesto, tres SP, repositorio, pruebas de alteración | Implementado |
| Una OC por presupuesto; números interno/cliente separados | Restricción única y SeguimientoOC_Guardar; pruebas SQL | Implementado |
| Una guía ligada al presupuesto; emergencia sin OC | Restricción única, SP, botón/popup y pruebas SQL/Chrome | Implementado |
| Confirmación desde OC, fecha/número/observación y archivo | Botón, formulario, SP y pruebas MVC/Chrome | Implementado |
| Una confirmación vigente e historial de retiradas | Índice filtrado, transición y pruebas SQL de reemplazo | Implementado |
| Aceptación guía con fecha/representante y estados independientes | CambiarEstado, fichas y pruebas SQL/HTTP/Chrome | Implementado |
| Ejecución, cierre y reapertura con dependencias | Cuatro SP CambiarEstado, acciones MVC y pruebas SQL | Implementado |
| Archivos privados visualizables/descargables y retiro auditado | SQL de autorización, rutas MVC, controles de formato/tamaño/hash | Implementado; pruebas locales con sustitutos |
| Adjuntos en R2 y bucket definido por entorno | R2DocumentFileStore, cliente S3 y .env.example | Preparado para hyt-documentos; bucket/claves los configura el usuario |
| Turnstile en acceso, sin omitir validación del servidor | Widget, Siteverify, acción/dominio/fecha, pruebas HTTP y de fallos | Implementado; claves reales pendientes del usuario |
| Configuración SQL/Cloudflare/Turnstile en User Secrets propio | UserSecretsId web y JSON privado solicitado; CLI .NET y SQL verificados | Creado y vinculado; conexión local válida |
| Auditoría visible, permisos, concurrencia y reintentos | Historial consolidado, tokens, DocumentoOperacion, pruebas SQL/HTTP | Implementado |
| Error de auditoría revierte estado | Trigger de fallo solo en base temporal y comprobación posterior | Verificado |
| Navegación desde popup a estados e historial y entre documentos | _Form, _Related, Context y pruebas HTTP/Chrome | Implementado |
| SQL ejecutado directamente en desarrollo | Aplicaciones sqlcmd verificadas; base local, fuentes y manifiestos | Aplicado; sin registros comerciales de prueba |
| Paquetes incrementales, orden y huellas | ETAPA_5/6_HYT.sql, SHA256_ETAPA_5/6.json; instalación aislada | 97 pruebas aprobadas desde paquetes |
| Documentación de estructura/SP | DOCUMENTACION_BD_Y_FUNCIONALIDADES.md y guía 5/6 | Actualizada |

Resultados consolidados en entregables/VERIFICACION_ETAPAS_5_Y_6.txt. Total actual: 142 fuentes SQL, 47 tablas y 105 relaciones verificadas offline; etapas 5/6 aportan 13 y 20 fuentes respectivamente.

La prueba real de R2/Turnstile requiere los recursos y claves que el usuario proporcionará después. No se simula como conexión real ni se crean buckets por la aplicación. El primer administrador se aprovisiona conforme a la guía de acceso. Marcado Facturado, dashboard y despliegue corresponden a las etapas 7, 8 y 9 del plan.