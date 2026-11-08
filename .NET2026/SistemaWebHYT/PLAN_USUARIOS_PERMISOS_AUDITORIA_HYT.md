# Plan de trabajo: usuarios, permisos y auditoría HYT

Fecha del plan: 06/11/2026. Implementación y verificación local: 07/11/2026.

## Objetivo y alcance

Administración sencilla para una sola empresa. El administrador da de alta usuarios directamente, sin autorregistro, aprobación adicional ni confirmación del correo. El correo se usa para iniciar sesión. Cada usuario recibe una contraseña temporal, debe reemplazarla antes de entrar al sistema y dispone siempre de Cambiar contraseña.

Se conservan MVC C#, ADO.NET y procedimientos almacenados. Se reutilizan AspNetUsers, AspNetRoles, AspNetUserRoles, Permiso, RolPermiso y EventoAuditoria. Ya existen hash de contraseñas, bloqueo por intentos, validación de sesiones y recarga de permisos. Los accesos de menú Usuarios y permisos/Auditoría requieren sus pantallas administrativas.

## Experiencia del administrador

Menú Administración con Usuarios y permisos (pestañas Usuarios y Perfiles de acceso) y Auditoría. Listados al ancho disponible, buscador, filtros sencillos, paginación SQL de 20, acciones uniformes, títulos y popups HYT, modo oscuro y espera durante operaciones.

### Usuarios

Listado: nombres y apellidos, correo, perfil, activo/inactivo, situación del acceso (Disponible, Bloqueado por intentos o Pendiente de cambiar contraseña), último ingreso y acciones. Activo/inactivo y obligación de cambio son propiedades distintas; no se mezclan con estados de documentos comerciales.

Nuevo usuario: Nombres, Apellidos, Correo, Perfil de acceso, Contraseña temporal y Activo. Correo único, sin distinguir mayúsculas/minúsculas. La contraseña temporal puede generarse y copiarse en ese momento; nunca se recupera después desde la base. Si se pierde, se genera otra mediante Restablecer contraseña.

Acciones según permiso: Ver, Editar datos/perfil, Activar/Desactivar, Restablecer contraseña, Desbloquear acceso y Cerrar sesiones. La ficha muestra Fechas y responsables en el estilo existente. No se elimina físicamente a usuarios con historial.

El administrador no puede consultar la contraseña actual de otro usuario. Desactivar, restablecer contraseña o cerrar sesiones invalida sus accesos anteriores. No se permite desactivar ni quitar privilegios al último administrador activo; se controla también con operaciones concurrentes en SQL.

### Perfiles de acceso

Se usa un perfil principal por usuario en la interfaz inicial, compatible con la relación existente de roles. El administrador puede crear, editar y duplicar perfiles, y consultar cuántos usuarios los utilizan. No se incorporan excepciones individuales por usuario en la primera entrega: se evita una mezcla de permisos difícil de explicar.

Perfiles iniciales sugeridos y ajustables:

| Perfil | Uso |
| --- | --- |
| Administrador | Usuarios, perfiles, configuración y acceso autorizado a todos los módulos. |
| Comercial | Clientes, tiendas, contactos, materiales/tarifarios y presupuestos. |
| Operaciones | Consulta de presupuestos y gestión de OC, Guía, ejecución y confirmación, según asignación. |
| Administración | Confirmación, referencia Facturado y reportes, según asignación. |
| Consulta | Solo lectura e impresión/descarga expresamente autorizada. |

Matriz con filas por módulo y casillas de acciones en lenguaje claro: Consultar, Crear, Editar, Activar/Desactivar, Cambiar estado, Imprimir y Descargar, únicamente cuando correspondan. Facturado y Restablecer contraseña son acciones identificables. Se pueden marcar todas las acciones de un módulo o copiar un perfil; se advierten dependencias y se incluye resumen antes de guardar en la misma pantalla, sin una aprobación separada.

Los códigos CONSULTAR/EDITAR actuales se reutilizan donde correspondan. Se agregan permisos más específicos cuando sea necesario separar acciones hoy agrupadas; no se muestran códigos internos al administrador. Ocultar un botón no sustituye validaciones en servidor y SQL. Configurar un perfil nunca elimina los requisitos comerciales: por ejemplo, Facturado sigue exigiendo confirmación vigente y guía.

## Experiencia del usuario

Flujo de primer ingreso:

1. El administrador registra el usuario y comunica sus credenciales temporales por el medio habitual de la empresa.
2. El usuario inicia sesión con correo y contraseña temporal; se mantiene Turnstile y sus mecanismos existentes.
3. Se presenta Cambia tu contraseña con Nueva contraseña y Confirmar nueva contraseña. Esa sesión solo permite cambiarla o cerrar sesión; URLs directas, peticiones de archivos y otras acciones permanecen bloqueadas.
4. Se comprueba que la clave nueva sea válida y diferente de la temporal. Se guarda su hash, se retira la obligación de cambio y se invalidan las sesiones anteriores.
5. Se renueva la sesión y se abre el inicio autorizado según su perfil.

**Cambiar contraseña** aparece en la cabecera para todos los perfiles, independiente de los permisos administrativos. En un cambio voluntario se solicita contraseña actual, nueva y confirmación; no requiere intervención del administrador. Tras restablecimiento administrativo se vuelve a exigir el cambio obligatorio. Cada campo se muestra en una fila y tiene un icono para mostrar u ocultar su contenido. El mismo control se usa al iniciar sesión y al registrar una contraseña temporal.

La contraseña temporal se muestra solo al crearla/restablecerla, con opción de mostrar/ocultar y copiar. No se guarda en texto claro ni se incluye en correos automáticos, auditorías, exportaciones o logs. La validación de fortaleza debe ser clara para el usuario y aplicarse en servidor. No se impone renovación periódica automática en este alcance.

## Auditoría amigable

Auditoría en una pantalla independiente, solo para perfiles autorizados. Filtros: rango de fechas (último mes por defecto), usuario, módulo, acción, registro y búsqueda. Listado paginado SQL de 20: fecha/hora de Perú, usuario, módulo, acción y referencia reconocible (número de presupuesto, OC, Guía o usuario afectado).

Ver detalle abre popup con responsable, fecha del evento, motivo cuando corresponda y tabla Campo / Antes / Después con etiquetas de negocio. No se muestra JSON como presentación principal. Un evento anterior con información parcial se presenta honestamente, sin inventar diferencias. Auditoría se consulta; no tiene botones de editar/eliminar. Se conserva la protección inmutable ya existente.

Eventos a cubrir: alta y edición de usuario, perfil asignado/cambiado, permisos modificados, activación/desactivación, contraseña temporal emitida/restablecida, cambio de contraseña obligatorio/voluntario, bloqueo/desbloqueo, cierre de sesiones, ingresos/cierres de sesión, cambios comerciales y estados ya auditados. Los eventos actuales se reutilizan; se completan huecos detectados en el inventario. Contraseñas, hashes, claves, cookies y tokens nunca forman parte del detalle. Se conserva el motivo de cambios cuando lo exige la regla vigente.

## Etapas de implementación y criterio de cierre

| Etapa | Entrega | Verificación |
| --- | --- | --- |
| 1. Inventario y diseño | Mapa de permisos actuales por acción/pantalla/SP, perfiles iniciales, pantallas y migración compatible. | Cada acción del sistema tiene permiso y regla de negocio identificados; se detectan permisos compartidos que deban separarse. |
| 2. Base SQL y contratos | Scripts idempotentes, campos y SP necesarios, contratos ADO.NET y documentación. | Correos únicos, hash seguro, concurrencia, auditoría atómica, ningún usuario existente pierde acceso por la migración. |
| 3. Usuarios y claves | Listado y ficha, alta directa, restablecimiento, bloqueo/cierre de sesiones y primer cambio obligatorio. | Usuario temporal no puede saltarse el cambio ni usando URLs/API/adjuntos; clave anterior deja de funcionar. |
| 4. Perfiles y aplicación de permisos | Matriz amigable, duplicar perfiles y asignar perfil; protección de acciones completas. | Menús/botones y acceso directo coinciden; retirar permiso afecta sesiones existentes; último administrador protegido. |
| 5. Mi cuenta | Cambiar contraseña para todos los usuarios y mensajes claros. | Se valida contraseña actual en cambios voluntarios; temporal no se reutiliza; sesiones anteriores revocadas. |
| 6. Auditoría | Buscador, filtros, referencias y detalle Antes/Después. | Acciones administrativas/comerciales visibles, sin secretos; auditoría inmutable y accesible solo con permiso. |
| 7. Verificación y entrega | Pruebas con perfiles reales, escritorio/móvil/oscuro, scripts y publicación preparada. | Pruebas de aceptación completas, SQL local validado en base aislada y entrega documentada. |

## Preparación técnica

AspNetUsers necesita separar Apellidos (hoy existe Nombres), registrar obligación de cambiar contraseña y fecha de cambio, conservar responsable de alta/actualización y permitir consultar último ingreso. Reutilizar datos de auditoría donde ya estén disponibles antes de duplicarlos. El correo normalizado necesita unicidad; antes de crear el índice se revisan duplicados y se informa cualquier conflicto sin modificar cuentas silenciosamente. El cambio de contraseña/permiso utiliza los sellos de seguridad y concurrencia ya existentes.

Procedimientos sugeridos por responsabilidad: listar/obtener/guardar usuario; activar/desactivar; restablecer contraseña; desbloquear; invalidar sesiones; cambiar contraseña propia; listar/obtener/guardar/duplicar perfiles; asignar perfil; consultar auditoría. Todos CREATE OR ALTER, NOCOUNT, columnas explícitas, TRY/CATCH, permisos y auditoría transaccional. La generación/verificación del hash permanece en C#, con el mecanismo existente; SQL persiste el hash y controla la operación autorizada, nunca recibe contraseñas en texto claro.

La migración no divide nombres completos existentes por heurísticas ni fuerza a cambiar contraseñas a todos sin necesidad. El administrador podrá completar apellidos y aplicar restablecimiento cuando corresponda. No se agrega multiempresa, registro público, confirmación de correo ni módulo de envío de invitaciones. No se modifica la configuración de R2/Turnstile por este trabajo.

## Pruebas de aceptación esenciales

- Crear usuario con correo único; rechazar duplicados incluso con diferencias de mayúsculas o altas simultáneas.
- Entrar con clave temporal, quedar restringido al cambio obligatorio y luego acceder a los módulos de su perfil.
- Cambiar la contraseña propia desde cualquier perfil; rechazar clave actual incorrecta y confirmación diferente.
- Restablecer clave: revocar accesos anteriores y exigir cambio nuevamente.
- Desactivar o retirar permisos: accesos existentes y peticiones directas dejan de funcionar.
- Impedir que dos operaciones concurrentes dejen la empresa sin administrador activo.
- Usuario de consulta no modifica datos aunque intente enviar una petición manual.
- Mantener restricciones comerciales de OC, confirmación, Guía y Facturado aunque el perfil tenga edición.
- Auditoría registra responsables y diferencias sin contraseñas/hash/tokens; filtros y paginación correctos.
- Formularios, mensajes, loading, modo oscuro y popups legibles en escritorio y móvil.

Orden recomendado: etapas 1–2 como base; 3 y 5 completan el ciclo de contraseña; 4 consolida permisos; 6 expone auditoría; 7 valida y entrega. Todo el alcance queda planificado antes de implementar.

## Entrega del 07/11/2026

Las pantallas están disponibles en Administración y el acceso a **Cambiar contraseña** está en la cabecera. Se aplicó el paquete de administración a la base local SistemaWebHYT y se comprobó que la cuenta administradora existente conserva acceso sin cambio obligatorio. Para otra instalación actualizada a las etapas previas, aplicar entregables/ADMINISTRACION_HYT.sql con SQLCMD y opción -b antes de instalar la versión web.

Compilación sin errores ni advertencias. Se aprobaron 309 comprobaciones HTTP y C#, 268 comprobaciones en una base SQL aislada instalada desde los paquetes y la validación de 183 fuentes SQL. Chrome comprobó las pantallas a 1440 y 390 píxeles, la selección de permisos, el detalle de auditoría y la persistencia del modo oscuro. La compilación Release quedó en entregables/publicacion.
