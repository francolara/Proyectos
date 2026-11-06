# HYT — acceso y presentación, etapa 2

03/11/2026. Código y scripts preparados. Ninguna conexión o ejecución SQL realizada por Codex.

## Despliegue manual

1. Revisar y ejecutar primero el paquete de etapa 1 en la instancia elegida por el usuario.
2. Revisar y ejecutar `entregables/ETAPA_2_HYT.sql` después de la etapa 1. Añade cinco procedimientos de acceso y un script de permisos, sin tablas nuevas ni cuentas predeterminadas.
3. Configurar la conexión **ConnectionStrings:HYT** en un almacén de secretos o variable de entorno **ConnectionStrings__HYT**. No guardarla en appsettings.json ni en Git. La cuenta web usa el rol SQL hyt_app_executor, sin db_owner y sin permisos de aprovisionamiento administrativo.
4. Crear el primer administrador mediante la herramienta manual descrita abajo, usando una conexión de despliegue autorizada. No existe usuario o contraseña de fábrica.
5. Para desarrollo iniciar el perfil **https** de GestionServicios.Web; las cookies solo funcionan con HTTPS. La revisión visual HTTP local no inicia sesión y no sirve como configuración de uso real.
6. En producción configurar **Security:DataProtectionPath** en un directorio privado persistente fuera de wwwroot y de la imagen del contenedor. Proteger acceso y claves en reposo conforme a la infraestructura; la aplicación se detiene si falta esta ubicación fuera de Development. Mantener el mismo directorio entre reinicios. Revisar HTTPS/proxy de confianza antes del despliegue siguiendo DOCUMENTACION_INFRAESTRUCTURA_FRALSE.md de la raíz.

El host nunca crea bases, ejecuta migraciones ni aplica los paquetes automáticamente. La aplicación puede mostrar el acceso sin conexión; un intento de ingreso sin configuración válida informa indisponibilidad y no permite entrar. No hay modo de autenticación simulada habilitable en producción.

## Herramienta administrativa — ejecución exclusiva del usuario

Proyecto: `tools/Hyt.Admin`. `--help` solo muestra instrucciones y no abre conexiones. Comandos reales **crear-admin** y **restablecer** sí se conectan cuando el usuario los ejecuta deliberadamente. Codex únicamente comprobó `--help`.

| Variable | Uso |
|---|---|
| ConnectionStrings__HYT | Conexión administrativa, distinta de la cuenta limitada de la web |
| HYT_ADMIN_USERNAME | Usuario de 3–128 letras sin tildes, números, punto, guion o guion bajo |
| HYT_ADMIN_PASSWORD | Contraseña introducida por el administrador, mínimo 12 caracteres con mayúscula, minúscula, número y símbolo |
| HYT_ADMIN_NAME | Nombre visible para crear el primer administrador |
| HYT_ACTOR_ID | Id del administrador activo con USUARIOS.EDITAR para restablecer acceso |
| HYT_RESET_REASON | Motivo del restablecimiento |

Usar las variables de forma temporal en el proceso o un mecanismo seguro de configuración; no pegarlas en chats, scripts versionados ni líneas de comando registradas. Quitar los valores del entorno al terminar. El programa genera el hash con PasswordHasher de ASP.NET Core Identity, nunca muestra contraseña o conexión, y registra auditoría. El alta inicial devuelve únicamente el identificador de la cuenta para la administración posterior. La recuperación de acceso de esta entrega es **asistida por el administrador**, no por correo automático ni autorregistro público. Comunicar la nueva contraseña por el canal privado que HYT autorice, fuera de esta herramienta.

El bootstrap se serializa en SQL y rechaza crear otro administrador inicial cuando ya existe uno activo. El restablecimiento cambia hash y sello, limpia bloqueo y conserva el estado activo/inactivo de la cuenta; no reactiva una cuenta por accidente. Los SP administrativos no están concedidos al rol web.

## Seguridad de acceso

- Autenticación mediante cookie de ASP.NET Core, hashes estándar Identity y repositorio ADO.NET + SP. No se utiliza EF, migración automática ni un IUserStore/SignInManager ficticio.
- Cookie __Host-Hyt.Session: Secure, HttpOnly, SameSite=Lax, 30 minutos de vigencia renovable y límite absoluto de ocho horas. La sesión no se persiste voluntariamente al cerrar el navegador.
- Relectura de usuario, sello y permisos en cada solicitud autenticada. Cuenta desactivada, bloqueo, cambio de sello o fallo de repositorio rechazan la sesión. La aplicación MVC no mantiene circuitos de componentes; cada solicitud relee identidad y permisos, y las futuras operaciones SQL seguirán validando permisos por escritura.
- Cinco contraseñas fallidas bloquean 15 minutos mediante el SP transaccional. El host limita solicitudes al acceso por dirección IP a 20/minuto. Detrás de proxy, configurar la política de proxies confiables antes de usar IP real; no confiar directamente en cabeceras enviadas por el cliente.
- Errores de credenciales genéricos; no revelar si existe o está desactivada una cuenta. Sin contraseñas, hashes o conexiones en logs o HTML.
- Formularios MVC con antifalsificación global; retorno tras login solo local. Salir mediante POST; GET solo pide confirmación. Inicio y cierre quedan auditados cuando el repositorio está disponible. Si falla la auditoría de salida, se cierra de todos modos la sesión local y queda advertencia sin datos sensibles.
- Políticas por claim `hyt:permission`; menú filtrado por permisos y controladores de módulo con comprobación de autorización antes de devolver la vista. Los módulos de negocio aún no tienen operaciones: al implementarlos deberán aplicar autorización de endpoint/servicio, además de la presentación de botones y las validaciones SQL.
- Respuestas HTML sin almacenamiento en caché. Datos internos de AccessAccount no se exponen como respuesta pública ni se deben serializar como DTO de usuario.

## Pantallas y componentes comunes

Acceso HYT en rojo/blanco/negro con logotipo original, mostrar/ocultar contraseña, errores de campo y resumen. Menú lateral en escritorio y botón plegable en celular; inicio y nombre de usuario, cierre de sesión, grupos Maestros, Comercial, Seguimiento, Consultas y Administración.

Los módulos de negocio se muestran como «en preparación» solo a cuentas autorizadas. Esta entrega no simula clientes, presupuestos ni indicadores. El dashboard conectado y las operaciones continúan en sus etapas del plan.

| Parcial MVC | Contrato |
|---|---|
| _Grid.cshtml | Hasta 20 filas de SQL, total long y enlaces con filtros; sin paginar un catálogo completo en navegador |
| _Filters.cshtml | Formulario GET con código/descripción y Desde/Hasta opcionales; valores iniciales del último mes en Perú |
| _RecordForm.cshtml | Nuevo/Editar/Ver; campos definidos por parcial del módulo; antifalsificación y validación definitiva en controlador |
| _Modal.cshtml | Dialog nativo, foco, Escape y bloqueo de cierre ocupado mediante hyt-ui.js |
| _DatePicker.cshtml | Campo date para model binding y calendario de FralseCont con presentación dd/MM/yyyy, Hoy/Limpiar |
| _Trace.cshtml | Autoría, fechas e historial, presentado en America/Lima |

Los modelos de presentación están en GestionServicios.Web/ViewModels/Shared y nunca se usan para ejecutar SQL desde una vista. El módulo construye nombres de parciales, controladores y acciones en código del servidor: no aceptarlos libremente de formularios o de datos no confiables. Los controladores POST deben validar ModelState, permisos, estado y concurrencia; el modo de consulta visual no sustituye esas comprobaciones.

La aplicación activa usa MVC exclusivamente, sin proyecto UI de Blazor ni Razor Pages. La futura APK compartirá reglas/API, no una biblioteca de componentes Razor. Colocar ayudas con formularios fuera del formulario principal para no anidar forms. Cada módulo conservará filtros en URL al regresar del detalle. El calendario está adaptado de FralseCont con colores HYT, fechas locales y etiquetas accesibles.

## Verificación

- Solución compilada, pruebas HTTP con WebApplicationFactory y repositorio exclusivo de pruebas (no un modo de producción), y pruebas HTTP de vistas parciales MVC.
- Chrome real: acceso 1440x1000 y 390x844 sin desbordamiento, mostrar/ocultar contraseña y módulo real del popup con foco/Escape/guardado ocupado. Capturas en entregables/preview.
- Sintaxis y dependencias SQL offline mediante ScriptDom; paquete incremental y huellas SHA256.
- Sin pruebas ejecutadas contra SQL Server. La transacción de bloqueo/auditoría y los permisos efectivos SQL requieren la revisión y ejecución manual del usuario en ambiente de pruebas.

Comandos reproducibles desde la raíz del repositorio (los de pruebas no conectan SQL):

```text
dotnet build .NET2026/SistemaWebHYT/SistemaWebHYT.slnx --nologo -m:1
dotnet run --project .NET2026/SistemaWebHYT/tools/Hyt.AccessTests
dotnet run --project .NET2026/SistemaWebHYT/tools/Hyt.SqlValidation -- Basededatos/Dbsishyt
dotnet run --project .NET2026/SistemaWebHYT/tools/Hyt.Admin -- --help
```

Para revisar el navegador iniciar temporalmente el host local y ejecutar tools/check_login_browser.mjs. Esa prueba solo abre el acceso y crea una ventana modal temporal para comprobar su módulo JS; no ingresa credenciales ni registra datos. Detener el host al terminar.

## User Secrets de desarrollo — 03/11/2026

Proyecto vinculado mediante UserSecretsId: `aspnet-SistemaWebHYT-0b5cb513-f035-40a5-b180-58a7e2a649a3`.

Archivo privado creado, siguiendo el formato de claves con dos puntos utilizado por FralseCont:

`C:\Users\Franco Lara\AppData\Roaming\Microsoft\UserSecrets\aspnet-SistemaWebHYT-0b5cb513-f035-40a5-b180-58a7e2a649a3\secrets.json`

Incluye ConnectionStrings:HYT configurada para la instancia local localhost\SQLEXPRESS, base SistemaWebHYT y autenticación integrada de Windows. En el mismo JSON están las claves CloudflareR2 y CloudflareTurnstile, con bucket hyt-documentos y dominios locales. Los valores privados de Cloudflare quedan vacíos para que el administrador los complete; no se copiaron credenciales de otros proyectos. La creación del bucket R2 sigue pendiente por indicación del usuario.

ASP.NET Core carga este archivo automáticamente en Development por el UserSecretsId del proyecto web. En producción se utilizan variables del proceso/contenedor. El JSON privado está fuera del repositorio. No se publican claves secretas en esta documentación y no se ha creado un administrador inicial.

## Captcha de respaldo y loading — 05/11/2026

La verificación de acceso admite Turnstile o captcha manual validado por el servidor. Ante falta de clave pública se muestra directamente el código manual; errores o demora del widget ofrecen el mismo respaldo, sin omitir contraseña ni protección del acceso. Se agregó indicador compartido para operaciones demoradas y protección de doble envío. Reglas, límites y pruebas en CAPTCHA_Y_LOADING.md. Las notas anteriores describen el estado de sus fechas; esta sección actualiza la política vigente de respaldo.
