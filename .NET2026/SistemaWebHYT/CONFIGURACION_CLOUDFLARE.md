# Configuración de Cloudflare para HYT

Fecha: 03/11/2026. Adjuntos en Cloudflare R2 y verificación de acceso con Cloudflare Turnstile.

El bucket propuesto es **hyt-documentos**. El nombre está en la variable `CloudflareR2__BucketName` del archivo `.env.example` y es también el valor predeterminado del sistema. El administrador creará el bucket; la aplicación no crea buckets.

| Variable de entorno | Valor o uso |
|---|---|
| `CloudflareR2__BucketName` | `hyt-documentos` |
| `CloudflareR2__AccountId` | ID de la cuenta de Cloudflare, de 32 caracteres hexadecimales |
| `CloudflareR2__AccessKeyId` | Access Key ID del token R2 |
| `CloudflareR2__SecretAccessKey` | Secret Access Key del token R2; solo servidor |
| `CloudflareTurnstile__Enabled` | `true`; mantiene activa la verificación de acceso |
| `CloudflareTurnstile__SiteKey` | Clave pública del widget de Turnstile |
| `CloudflareTurnstile__SecretKey` | Clave secreta de Turnstile; solo servidor |
| `CloudflareTurnstile__AllowedHostnames__0` | Dominio real de la web, sin protocolo ni ruta |

Para más dominios permitidos, añadir `CloudflareTurnstile__AllowedHostnames__1`, `__2`, etc. El ejemplo usa localhost, que se debe reemplazar al publicar. Registrar los mismos dominios en el widget de Cloudflare.

## Preparación de R2

1. Crear el bucket `hyt-documentos` y mantenerlo privado, sin habilitar acceso público r2.dev.
2. Crear credenciales S3 de R2 con lectura y escritura de objetos limitadas a este bucket.
3. Completar las cuatro variables R2. El endpoint se construye como `https://<AccountId>.r2.cloudflarestorage.com`; región de firma `auto`.
4. Cargar las variables en el proceso web o contenedor y reiniciar la aplicación.
5. Registrar una guía PDF, visualizarla, descargarla y volver a abrir el registro para confirmar su conservación.

Los archivos se envían desde el servidor a R2. El navegador utiliza las rutas autenticadas de HYT, por lo que no necesita acceso público ni CORS directo al bucket. SQL conserva el vínculo, nombre original, clave opaca, tipo, tamaño y SHA256; no almacena credenciales ni el contenido binario. Se mantienen las restricciones de PDF para guía, PDF/PNG/JPG para OC y confirmación, y 15 MB por archivo.

Los reintentos usan una escritura condicional que no sobrescribe objetos. Si el objeto ya existe se comprueba el contenido; cada lectura verifica tamaño y huella antes de entregarlo. Retirar un adjunto conserva el objeto y la auditoría y bloquea su acceso a través del sistema. Una operación SQL de resultado incierto no provoca borrado automático en R2.

## Preparación de Turnstile

1. Crear un widget de Turnstile para el dominio de HYT, con modo Managed.
2. Copiar Site Key y Secret Key a las variables indicadas.
3. Registrar los dominios autorizados también en las variables AllowedHostnames.
4. Reiniciar y verificar un inicio de sesión válido y otro sin completar el desafío.

La aplicación valida el token con Siteverify antes de consultar al usuario o verificar la contraseña. Comprueba éxito, acción `login`, dominio autorizado y fecha del desafío. Un token ausente, rechazado, vencido o una falla de comunicación no omite la verificación. La clave secreta nunca se incluye en la página.

Actualización 05/11/2026: en producción siempre se exige una verificación válida, mediante Turnstile o su captcha manual de respaldo. Enabled=false no elimina este requisito. Si Turnstile no está disponible, no se autentica automáticamente: se exige resolver el desafío manual antes de verificar la contraseña. Las pruebas aisladas desactivan explícitamente el desafío en Development o simulan las respuestas de Cloudflare; no utilizan claves reales. Si se configura una CSP en el proxy, permitir los recursos de Turnstile indicados por Cloudflare.

## Cómo cargar las variables

`.env.example` es una plantilla sin secretos. Copiarla a un archivo privado `.env` y completar los valores fuera del repositorio. Los archivos `.env` reales están ignorados por Git. ASP.NET Core no carga archivos `.env` automáticamente: usar variables de proceso, secretos locales o `env_file` en Docker Compose. No subir las claves a appsettings, SQL, Git ni documentos.

El proveedor de adjuntos registrado en la aplicación es R2. No hay respaldo automático a disco cuando faltan claves o falla Cloudflare. El almacenamiento local se utiliza únicamente al sustituir el servicio en pruebas aisladas.

## Estado de la verificación

Compilación y pruebas de contrato, integridad, idempotencia y acceso aprobadas con dobles de Cloudflare. Pruebas HTTP verifican que el token rechazado bloquea el acceso antes de consultar la cuenta y que el token válido permite continuar. No se creó ningún bucket ni se utilizaron credenciales reales. La prueba de conexión real queda para cuando el administrador cree el bucket y configure las claves.

Referencias oficiales: [R2 con SDK .NET](https://developers.cloudflare.com/r2/examples/aws/aws-sdk-net/), [escrituras condicionales R2](https://developers.cloudflare.com/r2/api/s3/extensions/), [validación Turnstile en servidor](https://developers.cloudflare.com/turnstile/get-started/server-side-validation/).
## User Secrets de desarrollo — 03/11/2026

Proyecto vinculado mediante UserSecretsId: `aspnet-SistemaWebHYT-0b5cb513-f035-40a5-b180-58a7e2a649a3`.

Archivo privado creado, siguiendo el formato de claves con dos puntos utilizado por FralseCont:

`C:\Users\Franco Lara\AppData\Roaming\Microsoft\UserSecrets\aspnet-SistemaWebHYT-0b5cb513-f035-40a5-b180-58a7e2a649a3\secrets.json`

Incluye ConnectionStrings:HYT configurada para la instancia local localhost\SQLEXPRESS, base SistemaWebHYT y autenticación integrada de Windows. En el mismo JSON están las claves CloudflareR2 y CloudflareTurnstile, con bucket hyt-documentos y dominios locales. Los valores privados de Cloudflare quedan vacíos para que el administrador los complete; no se copiaron credenciales de otros proyectos. La creación del bucket R2 sigue pendiente por indicación del usuario.

ASP.NET Core carga este archivo automáticamente en Development por el UserSecretsId del proyecto web. En producción se utilizan variables del proceso/contenedor. El JSON privado está fuera del repositorio. No se publican claves secretas en esta documentación y no se ha creado un administrador inicial.
## Respaldo manual y espera compartida — 05/11/2026

Por solicitud del usuario se incorpora el captcha manual de respaldo siguiendo el flujo de FralseCont. Aparece directamente si falta SiteKey; también puede elegirse, activarse al fallar el script/widget o si la verificación no se completa en doce segundos. Un rechazo del servidor ofrece el modo manual para el siguiente intento, sin autenticar por ese rechazo. Las claves Cloudflare siguen pendientes y no se modificaron.

El código tiene seis caracteres, cinco minutos de vigencia y un intento. Se guarda en memoria del servidor con límite de 5000 desafíos, ligado a una cookie segura HttpOnly del navegador; el HTML/JSON no contiene la solución. La imagen SVG privada no se almacena en caché. La aplicación actual usa una instancia; al reiniciarla se pide otro código. Una futura configuración con varias réplicas requerirá compartir el almacén de desafíos.

Los endpoints de imagen/renovación tienen un límite separado de 120 solicitudes/minuto por IP; los intentos de acceso conservan su límite de 20/minuto y bloqueo de cuenta. Resolver el captcha no sustituye contraseña, permisos, antifalsificación ni auditoría de acceso. Detalle de cobertura y pruebas en CAPTCHA_Y_LOADING.md.
