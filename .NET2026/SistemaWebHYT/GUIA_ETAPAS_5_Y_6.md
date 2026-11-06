# Guía de etapas 5 y 6 — HYT

Entrega de desarrollo: 03/11/2026. ASP.NET Core MVC .NET 10, C#, ADO.NET y procedimientos almacenados de SQL Server. Una empresa, identidad HYT roja, blanca y negra.

## Presupuestos

Abrir Presupuestos desde el menú. El listado inicia en el último mes, admite búsqueda y consulta veinte registros por página desde SQL. Cada presupuesto corresponde a un cliente y una tienda; atención por Invitación o Emergencia.

Crear un presupuesto, elegir cliente y tienda mediante las ayudas y agregar partidas del tarifario o manuales. Las zonas de material se muestran como información. La moneda inicial es PEN; USD es opcional y su precio se ingresa manualmente. Si falta precio, aparece pendiente y no se puede enviar. No existe conversión cambiaria automática.

Elegir CON IGV o SIN IGV. El sistema toma la tasa vigente a la fecha y conserva su valor por revisión. Cantidades, precios, conversiones unitarias e importes utilizan dos decimales. El servidor SQL calcula y valida; no acepta totales enviados por el navegador.

Guardar borrador; luego enviar y registrar la aprobación, observación, rechazo, vencimiento o cancelación permitidos, con fecha/responsable y motivo cuando corresponde. Una revisión enviada conserva su contenido; una corrección requiere nueva revisión. Solo se opera sobre la última revisión disponible. Con OC no se modifica contenido ni se crea otra revisión.

Desde la ficha generar Formato nuevo o Formato anterior. Ambos son HTML autónomos y pueden imprimirse desde el navegador. Cada generación conserva versión de plantilla, HTML, estilos, logo, usuario y huella de integridad. Consultar impresiones conservadas para recuperar exactamente el contenido generado; los cambios posteriores de maestros no lo reescriben.

## Invitación

1. Preparar, enviar y aprobar el presupuesto.
2. En el listado de presupuestos, abrir OC del cliente. Registrar su número externo alfanumérico y fechas. El sistema genera un número interno separado. Puede adjuntarse la OC.
3. Registrar ejecución y programación. Para iniciar se requiere OC activa, programación y responsable del trabajo.
4. Registrar guía/acta desde el botón del presupuesto: fecha, número alfanumérico, observación y PDF opcional.
5. Registrar aceptación de guía en su ficha, con fecha y representante del cliente. El registro de guía y su aceptación son acciones distintas.
6. Terminar ejecución con fecha real. Marcar trabajo completado en OC y cerrar operativamente cuando exista guía aceptada.

## Emergencia

Puede registrarse guía y ejecución sin OC previa; el trabajo puede iniciarse sin programación previa. El presupuesto sigue ligado a una sola tienda.

Cuando llegue la OC, vincularla a la revisión completa del presupuesto. Puede existir sin aprobación comercial previa. Después se permite regularizar la revisión vinculada: borrador → enviado → aprobado. Esta regularización no modifica precios, partidas ni importes.

Un presupuesto admite una OC y una guía. Cancelar o reactivar conserva los mismos registros, sujetos a sus dependencias y auditoría.

## Confirmación del cliente

Desde el listado de OC, usar Confirmación del cliente: fecha, número alfanumérico, representante, observación y adjunto opcional. El listado muestra su número, fecha y estado pendiente cuando aún no existe; permite filtrar confirmadas o sin confirmar.

La confirmación no cambia automáticamente la guía, la ejecución ni el estado comercial. Para retirar una confirmación abrir Ver registro, estados e historial y registrar motivo. Se conserva el antecedente y se puede registrar otra, con una sola vigente.

## Consulta y cambios de estado

Los popups de registros existentes ofrecen Ver registro, estados e historial. Las fichas enlazan presupuesto, OC, guía, ejecución y confirmación conforme a permisos. Un usuario de consulta puede abrir los documentos sin obtener permiso para modificarlos.

Las acciones de estado registran fecha de evento, responsable, motivo y estados anterior/nuevo. La grilla muestra registro, actualización y responsables; el expediente de presupuesto consolida auditoría paginada de la revisión seleccionada y documentos relacionados que el usuario puede consultar. Un fallo al auditar revierte la transición completa.

Reabrir requiere resolver dependencias: una guía aceptada no puede revocarse mientras OC permanezca cerrada; reabrir trabajo terminado exige reabrir OC y revisar guía aceptada. Cancelar OC exige retirar confirmación, anular guía y detener trabajo activo.

## Adjuntos y acceso

Los adjuntos reales se almacenan en Cloudflare R2, bucket previsto hyt-documentos. SQL conserva vínculo y metadatos. Guía acepta PDF; OC/confirmación PDF, PNG o JPG; máximo 15 MB. Visualizar y Descargar requieren permiso del documento. Retirar archivo exige motivo, conserva auditoría e impide su consulta posterior.

La configuración SQL local y los campos de Cloudflare están en el User Secrets del proyecto. Ver CONFIGURACION_CLOUDFLARE.md. La creación del bucket y carga de claves reales se dejan al usuario, como indicó. No se ha realizado una prueba contra su cuenta Cloudflare. Turnstile está integrado en login con validación en servidor; es obligatorio en producción.

Para ejecutar localmente, utilizar el perfil HTTPS del proyecto web. Falta aprovisionar el primer administrador conforme a GUIA_ACCESO_ETAPA_2.md; no existe usuario/contraseña predeterminados.

## SQL y evidencia

Las fuentes nuevas y modificadas están aplicadas en localhost\SQLEXPRESS, base SistemaWebHYT. Los paquetes ETAPA_5_HYT.sql y ETAPA_6_HYT.sql y sus huellas están en entregables. Son incrementales y requieren las etapas anteriores, en orden. No volver a aplicarlos para crear datos de prueba en la base de trabajo.

Verificación: 164 comprobaciones HTTP/C#, 97 comprobaciones SQL reales instalando los paquetes en base temporal eliminada; revisión Chrome de presupuestos, seguimiento, adjuntos, aceptación, regularización, filtros y móvil. Siete casos de formatos HTML/impresión verificados; los PDF de preview son evidencia de pruebas, no el motor del producto.

Las etapas siguientes del plan son Marcado Facturado (7), dashboard/reportes (8) y validación/despliegue (9). No se implementa emisión electrónica de comprobantes en esta entrega.