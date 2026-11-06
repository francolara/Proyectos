# Entrega SQL — etapa 1 de HYT

Fecha: 03/11/2026. Base: **SistemaWebHYT**. Requiere SQL Server 2022 o superior.

## Alcance y ejecución

Este paquete prepara una base nueva: 46 tablas, catálogos internos, roles/permisos, relaciones, auditoría y cuatro procedimientos de soporte. **Codex no se conecta ni ejecuta estos scripts.** No se crean cuentas SQL ni usuarios de aplicación, no hay contraseñas predeterminadas y no se copian datos de FralseCont.

El archivo consolidado está en `.NET2026/SistemaWebHYT/entregables/ETAPA_1_HYT.sql`, relativo a la raíz del repositorio. Contiene las mismas fuentes del manifiesto en orden y puede abrirse en SSMS para revisión y ejecución manual por el usuario. Alternativamente ejecutar cada fuente en el orden de la lista inferior. No ejecutar ambas opciones como pasos diferentes necesarios.

Antes de ejecutar: confirmar instancia y base destino; revisar el paquete; usar una cuenta de despliegue que pueda crear base, tablas y procedimientos. El nombre es SistemaWebHYT aunque la carpeta de fuentes se llame Dbsishyt. Los GO son separadores del cliente SQL, no instrucciones para enviar sin separar a SqlCommand.

Los scripts son repetibles sobre el mismo esquema inicial: crean objetos o filas si faltan y actualizan definiciones de SP/triggers. **No son un migrador de esquemas antiguos:** si ya existen tablas diferentes, detenerse y revisar diferencias; no borrar ni recrear bases de otros sistemas. Un error interrumpe el procedimiento de despliegue: revisar y resolver antes de continuar. El paquete completo no es una transacción única porque CREATE DATABASE requiere su propio contexto; las cargas se protegen con transacción.

## Datos iniciales

- Moneda: PEN / Soles / S/; USD / Dólares / $. Sin Id numérico.
- TipoDocumento: 01 Factura, 03 Boleta. Sin Id numérico.
- TipoIGV: CON_IGV / CON IGV y SIN_IGV / SIN IGV; no define tasas.
- TasaIGV: 0.18 desde 03/11/2026 como fecha de configuración del proyecto. Esta fecha no afirma inicio legal. Antecedentes anteriores necesitan vigencia validada antes de su carga.
- Roles: Administrador, Comercial, Operaciones, Seguimiento documental y Consulta; permisos iniciales propuestos, sin asignar usuarios.
- Numeración interna inicial PPTO- y OC-, contador 0, editable por despliegue antes de operar; numeración definitiva por confirmar. La OC externa nunca usa ese contador.
- No se importan el tarifario, clientes, materiales o contactos de los Excel; eso se hará con revisión de datos en etapas posteriores.

## Seguridad e integración posterior

Las tablas AspNet* usan claves string de 128 caracteres. Al implementar Identity configurar igual longitud y campos Nombres/Activo; acceso mediante stores ADO.NET + SP, sin migraciones EF automáticas. Los hashes los generará Identity; nunca insertar contraseñas en texto claro. El primer administrador se aprovisionará de forma segura al implementar la etapa 2, sin credenciales en Git.

El rol SQL hyt_app_executor no tiene escrituras directas sobre dbo: se ejecutan operaciones por procedimientos concedidos individualmente. El usuario de conexión y su asociación al rol los administrará el usuario; no usar db_owner ni sa en la aplicación. No hay GRANT general EXECUTE ni permiso público para el helper de auditoría. La conexión SQL no se publica ni se configura en este paquete.

La fábrica ADO.NET no abre conexiones por sí sola. En el host futuro, obtener UsuarioId de la identidad autenticada y no del formulario; cada operación de escritura se implementará con permisos, validación, transacción y llamada al helper de auditoría. Los procedimientos CRUD de todos los módulos, las pantallas y el login son entregables posteriores, no están simulados en esta etapa.

spu_RegistrarAuditoria exige transacción del llamador; un fallo debe revertir toda la operación de negocio. spu_RegistrarTasaIGV administra su propia transacción, cambia vigencia y registra auditoría. spu_ObtenerNumeroInterno reserva en la misma transacción del documento. spu_ListarClientes demuestra consulta por código/nombre con 20 resultados, total separado y control de permiso. Las demás búsquedas se implementarán al desarrollar cada módulo.

El trigger de auditoría impide UPDATE/DELETE. La protección de tasas impide solapamientos, y la clave compuesta en revisión conserva la tasa aplicada. El marcado Facturado requiere confirmación vigente, guía aceptada y trabajo terminado. Los retiros/correcciones se implementarán con permisos, motivo e historial en su etapa. Un administrador SQL privilegiado conserva sus poderes; estos controles no pretenden impedir administración legítima.

## Validación realizada y pendiente

La validación offline comprueba sintaxis mediante Microsoft ScriptDom, manifiesto, firmas, dependencias, columnas y claves candidatas de FK, convenciones de SP y restricciones principales. La compilación .NET comprueba los proyectos y referencias. No equivale a ejecutar SQL Server ni prueba concurrencia real, permisos efectivos o despliegue. Tras ejecución manual, se deberán comprobar esquema, cargas y operaciones en un ambiente de prueba antes de producción.

No hay recuperación de errores ni despliegue automático al iniciar la web. Los archivos de SHA256 permiten cotejar el paquete con las fuentes preparadas; cualquier modificación exige validar y empaquetar de nuevo.

## Orden del manifiesto

1. `Script/001_CrearBaseDatos.sql`
2. `Tablas/Moneda.sql`
3. `Tablas/TipoDocumento.sql`
4. `Tablas/TipoIGV.sql`
5. `Tablas/TasaIGV.sql`
6. `Tablas/AspNetUsers.sql`
7. `Tablas/AspNetRoles.sql`
8. `Tablas/AspNetUserRoles.sql`
9. `Tablas/AspNetUserClaims.sql`
10. `Tablas/AspNetRoleClaims.sql`
11. `Tablas/AspNetUserLogins.sql`
12. `Tablas/AspNetUserTokens.sql`
13. `Tablas/Permiso.sql`
14. `Tablas/RolPermiso.sql`
15. `Tablas/EventoAuditoria.sql`
16. `Tablas/ConfiguracionHYT.sql`
17. `Tablas/UnidadMedida.sql`
18. `Tablas/Zona.sql`
19. `Tablas/Especialidad.sql`
20. `Tablas/TipoServicio.sql`
21. `Tablas/Cliente.sql`
22. `Tablas/Tienda.sql`
23. `Tablas/ContactoCliente.sql`
24. `Tablas/ProveedorCliente.sql`
25. `Tablas/ServicioMaterial.sql`
26. `Tablas/MaterialZona.sql`
27. `Tablas/Tarifario.sql`
28. `Tablas/TarifarioVersion.sql`
29. `Tablas/TarifarioVersionZona.sql`
30. `Tablas/TarifarioPartida.sql`
31. `Tablas/TarifarioPartidaZona.sql`
32. `Tablas/Presupuesto.sql`
33. `Tablas/PresupuestoRevision.sql`
34. `Tablas/PresupuestoPartida.sql`
35. `Tablas/OrdenCompraCliente.sql`
36. `Tablas/Ejecucion.sql`
37. `Tablas/GuiaConformidad.sql`
38. `Tablas/GuiaPartida.sql`
39. `Tablas/SeguimientoCliente.sql`
40. `Tablas/ConfirmacionFacturacion.sql`
41. `Tablas/ReferenciaComprobanteExterno.sql`
42. `Tablas/ArchivoAdjunto.sql`
43. `Tablas/OrdenCompraArchivo.sql`
44. `Tablas/GuiaArchivo.sql`
45. `Tablas/ConfirmacionArchivo.sql`
46. `Tablas/ImpresionPresupuesto.sql`
47. `Tablas/NumeracionInterna.sql`
48. `Script/002_CargarTiposDocumento.sql`
49. `Script/003_CargarTiposIGV.sql`
50. `Script/004_CargarTasaIGV.sql`
51. `Script/005_CargarMonedas.sql`
52. `Script/006_CargarRolesPermisos.sql`
53. `Script/009_CargarNumeracion.sql`
54. `StoreProcedure/spu_ObtenerNumeroInterno.sql`
55. `StoreProcedure/spu_RegistrarAuditoria.sql`
56. `StoreProcedure/spu_ListarClientes.sql`
57. `StoreProcedure/spu_RegistrarTasaIGV.sql`
58. `Script/007_ProteccionAuditoria.sql`
59. `Script/008_PermisosAplicacion.sql`

## Entrega incremental de etapa 2

03/11/2026: manifest-etapa1.json conserva el orden original de esta guía; manifest.json es el acumulado de 65 scripts. Después de etapa 1, aplicar entregables/ETAPA_2_HYT.sql de SistemaWebHYT, o los seis archivos de manifest-etapa2.json en orden. No ejecutar el manifiesto acumulado como requisito para actualizar una base que ya tiene etapa 1.

La guía GUIA_ACCESO_ETAPA_2.md explica configuración HTTPS, secretos externos y aprovisionamiento manual del primer administrador. Las herramientas administrativas no se ejecutan desde el inicio de la web ni las ha ejecutado Codex contra un servidor. Nuevos SP: spu_ConsultarAcceso, spu_FinalizarAcceso, spu_RegistrarSalida, spu_CrearAdministradorInicial y spu_RestablecerAcceso. Solo los tres primeros se conceden al rol web. Los paquetes no crean logins SQL ni cuentas de usuario predeterminadas.

## Entregas incrementales de etapas 3 y 4

03/11/2026: manifiesto acumulado de 109 fuentes. Aplicar los paquetes ETAPA_3_HYT.sql (30 fuentes) y ETAPA_4_HYT.sql (14 fuentes) en ese orden después de etapas 1 y 2. Los manifiestos incrementales ofrecen las fuentes separadas en su orden. Las huellas SHA256 se entregan junto a cada paquete.

Etapa 3 agrega Activo a ProveedorCliente/TarifarioPartida y los procedimientos de mantenimientos/catálogos. Etapa 4 agrega versiones/partidas de tarifas, publicación, copia y vigencias IGV. Los permisos 012/013 conceden EXECUTE individual al rol web, sin habilitar DML directo. Detalle funcional y pruebas manuales en .NET2026/SistemaWebHYT/GUIA_ETAPAS_3_Y_4.md. No se ejecutó ningún paquete contra SQL Server desde Codex.
