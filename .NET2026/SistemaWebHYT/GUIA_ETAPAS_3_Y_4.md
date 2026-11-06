# Guía de entrega — mantenimientos y tarifarios

03/11/2026. Aplicación ASP.NET Core MVC .NET 10, ADO.NET y SP. Una sola empresa HYT. Scripts preparados, sin ejecución SQL por Codex.

## Orden manual

1. Aplicar primero los paquetes de etapas 1 y 2 en la base de pruebas SistemaWebHYT, si todavía no están aplicados.
2. Revisar y ejecutar `entregables/ETAPA_3_HYT.sql` con la cuenta de despliegue. Alternativa: fuentes de `Basededatos/Dbsishyt/manifest-etapa3.json` en orden.
3. Revisar y ejecutar `entregables/ETAPA_4_HYT.sql`. Alternativa: fuentes de `manifest-etapa4.json` en orden.
4. Mantener la conexión HYT en secretos externos. El usuario web usa el rol `hyt_app_executor`; los paquetes conceden EXECUTE a los SP de esta entrega. No crean logins ni usuarios predeterminados.
5. Mantener las instrucciones de infraestructura de .NET2026 y la guía de acceso de etapa 2. No hay cambio de despliegue ejecutado.

No es necesario ejecutar el manifiesto acumulado para actualizar una base que ya tiene etapas anteriores. Cada procedimiento usa CREATE OR ALTER. El script 011 agrega Activo a ProveedorCliente y TarifarioPartida solo si falta. Conservar respaldo y revisar la salida de ejecución antes de abrir los módulos.

## Navegación

Clientes abre el maestro; su ficha permite llegar a tiendas, contactos y datos de HYT por cliente, además del portal externo de documentos. Materiales ofrece accesos a especialidades y tipos de servicio. Unidades y zonas tienen sus propios mantenimientos. Tarifarios abre cabeceras; su ficha lleva a versiones y cada versión lleva a sus partidas. Configuración HYT abre vigencias del IGV.

Los mantenimientos se desactivan desde la ficha: no se eliminan sus antecedentes. Los permisos CONSULTAR y EDITAR se comprueban en servidor y en SQL. Consultas de referencia requieren también permiso de consulta del catálogo correspondiente. Las monedas y tipos IGV se consultan desde tablas internas; no tienen pantalla de edición.

## Comprobación manual contra SQL Server

- Crear cliente jurídico y natural, tienda y contacto. Comprobar que no acepta tienda de otro cliente; recuperar enlace documental sin credenciales.
- Crear unidad y zonas Sur/Centro. Registrar material con códigos `0001-A`/`000SAP` y ambas zonas. Abrir de nuevo y comprobar sus nombres y asociaciones.
- Buscar códigos/descripciones y recorrer más de veinte resultados; comprobar total y filtros al volver de la ficha.
- Desactivar una zona o unidad usada: conservar referencias históricas y no ofrecerla como nueva referencia. Editar otro campo del material sin perder la zona histórica.
- Abrir el mismo registro desde dos sesiones: guardar en una y comprobar rechazo de la otra, sin sobrescribir los cambios y conservando los datos introducidos en pantalla.
- Crear tarifario y versión. Registrar moneda PEN, tipo IGV y varias zonas predeterminadas. Crear partidas y comprobar los valores copiados; cambiar la cabecera y confirmar que las partidas anteriores no cambian.
- Dejar precio USD vacío: debe recuperarse vacío/NULL. Agregar USD manual y comprobar sus decimales; no debe existir conversión de moneda automática.
- Publicar versión con al menos una partida activa, PEN informado y material/unidad activos. Comprobar que versión y partidas publicadas no admiten edición, incluso con un POST directo.
- Copiar versión: nuevo número y estado BORRADOR, mismo contenido/precios/zonas, auditoría en origen/destino/partidas y contenido fuente intacto. Probar copia/publicación mientras otra sesión guarda una partida; el rowversion debe detectar el conflicto o las operaciones deben serializarse.
- Agregar vigencia IGV posterior con motivo; comprobar cierre exclusivo de la anterior, actor/fecha de ambos eventos y rechazo desde un formulario que conserva una última fecha desactualizada.
- Revisar EventoAuditoria y los historiales de ficha. Confirmar que usuario sin EDITAR no puede guardar y que la cuenta web no tiene DML directo.

El tipo IGV indica si el precio incluye impuesto; la tasa queda en TasaIGV por vigencia. El cálculo de presupuestos y su snapshot de tasa corresponde a etapa 5. Las versiones publicadas no cambian precios anteriores ni ejecutan actualizaciones de presupuestos.

## Validación local repetible, sin SQL

Desde la raíz del repositorio:

- `dotnet build .NET2026/SistemaWebHYT/SistemaWebHYT.slnx --no-restore`
- `dotnet run --project .NET2026/SistemaWebHYT/tools/Hyt.AccessTests --no-restore`
- `dotnet run --project .NET2026/SistemaWebHYT/tools/Hyt.SqlValidation --no-restore`
- `python .NET2026/SistemaWebHYT/tools/test_sql_validation.py`

Para Chrome, iniciar temporalmente `Hyt.AccessTests -- --browser-host` y ejecutar `node .NET2026/SistemaWebHYT/tools/check_maintenance_browser.mjs`. El host solo usa repositorios de prueba, escucha en loopback y elimina su archivo temporal al pulsar Enter. No incorporar los repositorios ni controladores de prueba al producto. Los archivos nuevos .cs/.cshtml y estáticos se incluyen mediante las convenciones del SDK, comprobadas por compilación y pruebas.

## Instalación local realizada por nueva autorización

03/11/2026: el usuario pidió crear la base local. Se ejecutaron las 109 fuentes en localhost\SQLEXPRESS, base SistemaWebHYT, con Windows integrado y QUOTED_IDENTIFIER habilitado (`sqlcmd -I`). Resultado: 46 tablas, 50 SP y 103 relaciones confiables/habilitadas. Refresco de módulos y comprobación de restricciones sin errores. La autorización anterior de solo preparar scripts queda sustituida para esta operación local; no se ha realizado despliegue de producción, configuración de secretos web ni alta del administrador.
