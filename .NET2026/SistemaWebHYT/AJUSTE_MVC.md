# Arquitectura definitiva HYT: MVC + ADO.NET + SP

03/11/2026. Cambio solicitado explícitamente por el usuario.

## Estructura activa

- GestionServicios.Web: Controllers, ViewModels, Views y wwwroot. ASP.NET Core MVC .NET 10.
- GestionServicios.Application y Domain: contratos de servicios y reglas de negocio C#.
- GestionServicios.Infrastructure: repositorios ADO.NET que ejecutan procedimientos almacenados.
- GestionServicios.Contracts: contratos de datos y paginación de 20 registros.
- Herramientas administrativas y de pruebas: siguen fuera de los flujos del producto.

La solución no referencia GestionServicios.UI, bUnit, Razor Components ni servicios/circuitos Blazor. Las pantallas de acceso dejaron de ser Razor Pages y son acciones MVC con vistas .cshtml. El renderizado Razor de .cshtml es parte de MVC y no implica usar Razor Pages o Blazor.

Controladores Access, Home y Module; vistas Access/Login, Access/Logout, Home/Index, Module/Index y errores compartidos. Rutas públicas conservadas: /Acceso, /Salir, /Denegado; inicio privado / y módulos /modulo/{code}. La autorización de módulos se comprueba en controlador y devuelve 403 si falta permiso.

Parciales reutilizables: _Grid, _Filters, _DatePicker, _RecordForm, _Modal y _Trace. Modelos tipados de presentación, HTML codificado, formularios MVC con antifalsificación global, GET para búsqueda, controladores POST para guardar y validar ModelState. Ninguna vista consulta SQL.

Calendario adaptado del JavaScript/CSS de FralseCont, con dd/MM/yyyy, valor ISO de formulario, lunes inicial, Hoy/Limpiar, límites y colores HYT. Popups con dialog nativo y JavaScript, foco/retorno y bloqueo ocupado. Se conserva presentación adaptable y logotipo HYT.

## Datos y seguridad conservados

Sin cambios SQL: se conservan todos los scripts de etapas 1/2, hashes, bloqueo de intentos, cookies seguras, sello de seguridad, permisos, auditoría y herramienta administrativa. No se ejecutó SQL Server ni se crearon usuarios reales. Los repositorios siguen usando CommandType.StoredProcedure, sin Entity Framework ni migraciones automáticas.

La implementación previa de Blazor se retiró de la solución y se conservó como respaldo local en tmp/hyt-blazor-pre-mvc-20261103, fuera del proyecto activo y del versionado. No forma parte del producto ni del despliegue.

## Verificación

Compilación limpia; pruebas HTTP de acceso y vistas parciales, comprobación de que todas las acciones son MVC y las páginas no cargan blazor.web.js. Chrome: acceso escritorio/móvil sin desbordamiento, popup y calendario MVC conservando fecha visible y valor de formulario. Pruebas contra motor SQL pendientes de ejecución manual del usuario.

Para una futura APK se reutilizarán servicios y reglas mediante API. Las vistas MVC no se convierten automáticamente en componentes MAUI; este cambio prioriza la tecnología pedida por el usuario antes de iniciar los módulos de negocio.
