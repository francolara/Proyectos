# Etapa 1: base técnica y SQL

Fecha: 03/11/2026. Estado: **preparación y validación offline completadas**. No se conectó ni ejecutó SQL Server.

## Entregables

- SistemaWebHYT.slnx: seis proyectos de producto (Domain, Contracts, Application, Infrastructure, UI y Web) y herramienta de validación offline.
- Contrato de página fija de 20 registros, códigos de negocio textuales y contexto de auditoría. Fábrica ADO.NET sin apertura automática de conexiones.
- 46 tablas y 103 relaciones, con una tienda por presupuesto, una OC y una guía como máximo por presupuesto, monedas/tipos como códigos textuales y precios USD opcionales.
- Esquema de Identity, roles y permisos, sin cuentas ni contraseñas predeterminadas. Conexión SQL mediante rol de ejecución limitada.
- Catálogos de moneda, documento, tratamiento IGV, tasa inicial versionada, roles/permisos y numeración interna. No se importaron los Excel.
- Cuatro SP base: auditoría transaccional, numeración interna, consulta paginada de clientes y nueva vigencia IGV auditada.
- Protecciones de auditoría, solapamiento de vigencias y requisitos del marcado Facturado.
- 59 scripts individuales en Basededatos/Dbsishyt, ordenados en manifest.json; guía LEEME_DESPLIEGUE.md y diccionario físico completo en DICCIONARIO_BD_ETAPA_1.md.
- Paquete entregables/ETAPA_1_HYT.sql y huellas SHA256 de fuentes/paquete para ejecución manual posterior.

## Evidencia verificada

- Compilación de la solución: cero errores y cero advertencias.
- Microsoft ScriptDom: sintaxis SQL Server 2022 válida, dependencias, columnas y claves candidatas de FK, firmas únicas, convenciones de SP y restricciones principales comprobadas offline.
- Cuatro pruebas negativas en copias temporales: rechaza sintaxis inválida, tabla FK inexistente, eliminación de unicidad de guía y SELECT * en SP. No tocan SQL Server.
- Fuentes nuevas y modificadas de la etapa revisadas en UTF-8 con CRLF. AGENTS.md original preservado.
- Paquete consolidado analizado y hashes cotejados con las fuentes.

## Límites y siguientes etapas

No se ejecutó el despliegue ni pruebas de concurrencia/permisos contra un motor SQL Server, por instrucción expresa del usuario. El paquete requiere revisión y ejecución manual en pruebas antes de producción; la validación offline no sustituye esas comprobaciones.

Esta etapa no incluye login operativo, CRUD completo, importación del tarifario, cálculo/impresión de presupuestos, archivos funcionales o dashboard conectado. Esos módulos continúan según el plan, empezando por etapa 2. Las tablas y helpers no implican que todas sus operaciones ya tengan procedimientos ni endpoints implementados.

Identity futuro usará stores ADO.NET con claves de 128 caracteres coherentes con el esquema; no activar migraciones automáticas. Regla de redondeo y semántica de los dos números del formato de presupuesto se concretarán antes de los cálculos/formularios, sin bloquear esta entrega estructural.

Actualización por cambio a MVC solicitado el 03/11/2026: la solución activa ahora tiene cinco proyectos de producto (Domain, Contracts, Application, Infrastructure y Web), más herramientas. La biblioteca UI de componentes se sustituyó por vistas y parciales MVC en Web; los scripts SQL de etapa 1 se conservan sin cambios.
