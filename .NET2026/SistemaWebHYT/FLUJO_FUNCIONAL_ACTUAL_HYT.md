# Flujo funcional vigente de HYT

Actualizado: 07/11/2026. Este documento consolida las decisiones vigentes y reemplaza las descripciones del flujo inicial cuando difieran. Sistema a medida de una sola empresa, MVC en C#, ADO.NET y procedimientos almacenados SQL Server.

## Presupuesto

Se registra para un cliente y una tienda. El número visible usa año/mes y correlativo mensual; la clave interna se conserva. Tipo de atención: Invitación o Emergencia. Un presupuesto puede tener como máximo una OC del cliente y una Guía; cada OC pertenece a un presupuesto. Los datos comerciales, moneda PEN/USD, importes con dos decimales y detalle sin IGV se mantienen. Los precios con IGV del tarifario se convierten a netos.

| Estado visible | Cuándo corresponde |
| --- | --- |
| Registrado | Al guardar el presupuesto; código interno BORRADOR. |
| Enviado | Al registrar el envío desde el listado. |
| Observado | Solicita motivo y permite corregir el presupuesto. |
| Aceptada | Aceptación manual o automática al registrar OC o Guía válidas. |
| Rechazada | El cliente rechaza; requiere motivo. |
| Cancelado | Se cancela con motivo y validación de documentos vinculados. |

Vencido no se utiliza. Los estados se gestionan desde el listado, no dentro de la ficha. El tipo Emergencia permite registrar la Guía sin OC previa; ambas se vinculan directamente al presupuesto. La aplicación valida contenido, permisos y concurrencia al guardar.

## OC del cliente

Conserva número interno y número externo alfanumérico, fechas de emisión/recepción, importe, observación y adjuntos. El cliente confirma mediante número alfanumérico, fecha, representante opcional, observación y adjuntos. Solo hay una confirmación vigente; las retiradas conservan su historial.

| Estado visible | Regla vigente |
| --- | --- |
| Registrada | Al registrar la OC. |
| Completado / Trabajo completado | Guardar una confirmación vigente completa automáticamente una OC registrada; el cambio manual también admite confirmación vigente o ejecución terminada. |
| Cerrada operativamente | Cierre manual con requisitos operativos validados: guía y ejecución terminada. |
| Cancelada | Cancelación con motivo; se deben resolver las dependencias vigentes. |

Editar una confirmación no degrada una OC cerrada. El cambio automático registra fecha, usuario, motivo, estados anterior/nuevo y referencia de confirmación en la misma transacción. Reintentar no duplica operaciones. No se modifica masivamente el estado de documentos históricos.

## Guía y ejecución

Ver y Editar de Guía comparten el mismo pop-up. Ver bloquea los campos y ofrece Editar registro según permiso; Editar ofrece Guardar registro. Desde Presupuestos se usa la misma ventana. La Guía tiene número alfanumérico, fecha, observación y PDF; no se administra un flujo de estados ni un catálogo de tipos adicional. En reportes el encabezado de la columna es **Guía**, tanto en el listado como en el HTML para imprimir.

Ejecución usa su ficha de responsable, fecha programada, inicio/término real y descripción. Desde Presupuestos se abre en pop-up para alta o edición, sin la antigua opción de cambio de estado. Los estados internos históricos se conservan por compatibilidad; no son un paso adicional del flujo visible. No se crean ni terminan ejecuciones automáticamente al confirmar una OC.

## Marcado Facturado

Es una referencia externa; no emite comprobantes. Requiere confirmación vigente del cliente de esa OC, guía registrada y trabajo completado acreditado por OC en Trabajo completado/Cerrada operativamente, o por ejecución terminada para compatibilidad con registros existentes. Una OC completada por confirmación no necesita una ejecución separada para este marcado.

Sin confirmación vigente se bloquea el guardado en SQL y se informa el requisito en la ficha. Se registra fecha del comprobante, tipo 01 Factura o 03 Boleta, serie y número conservando ceros. Correcciones/retiros exigen motivo y auditoría. La confirmación de un expediente marcado no puede retirarse sin retirar primero el marcado.

## Reportes integrales

El selector ofrece únicamente Presupuesto y Orden de Compra - Cliente. Se eliminan filtros separados Cliente/Tienda y Seguimiento de OC. Buscar incluye número de presupuesto, cliente, tienda y referencias documentales autorizadas. Cada fuente tiene sus estados propios; cambiar el reporte reinicia el estado dependiente.

- Presupuesto: estado, fechas de registro, búsqueda, tipo Invitación/Emergencia y checks Pendiente de guía/Pendiente de OC del cliente.
- OC cliente: sus cuatro estados, fechas de registro, búsqueda, tipo y checks Pendiente de guía, Pendiente de confirmación y Pendiente de facturar.
- Las fechas delimitan el registro del documento principal. Las relaciones se muestran aunque fueran registradas fuera del período, para no perder el expediente completo.
- Una fila muestra presupuesto, cliente/tienda, tipo, estado e importe; OC interna/externa y su estado; referencia de Guía; referencia interna de Ejecución; confirmación vigente; tipo/serie/número de comprobante vigente. Cada relación muestra su fecha de registro; las fichas conservan la auditoría completa.
- Los checks se combinan mediante AND. Pendiente de OC/Guía significa ausencia de documento registrado; no equivale a un estado. Pendientes de confirmación/facturar excluyen OC canceladas. Referencias retiradas no se presentan como vigentes.
- SQL pagina de 20 en 20 y calcula total e importes de toda la consulta. PEN y USD se muestran separados, sin sumarse ni convertirlos automáticamente. Los importes ausentes se presentan pendientes.
- Imprimir reporte completo abre un HTML independiente con logo HYT, filtros aplicados y todos los resultados, no solo la página visible. Formato A4 horizontal, encabezado de tabla repetido y opción de imprimir/guardar PDF desde el navegador.
- Los reportes y la impresión respetan permisos de consulta; los datos de relaciones sin permiso no se exponen. Los parámetros anteriores de los procedimientos se conservan solo por compatibilidad de integraciones, no como opciones de la pantalla.

## Inicio y dashboard

El inicio presenta el flujo vigente, tarjetas de todos los estados de presupuesto y OC con acceso al reporte correspondiente, y pendientes del período: OC, Guía, confirmación y facturación. Incluye listas para marcar Facturado y marcados vigentes. Los accesos de pendientes conservan fechas y filtros equivalentes a los contadores.

Se conservan curvas interactivas de registros por día/semana/mes; presupuestos y OC tienen gráficos de columnas independientes con sus etiquetas actuales. La dona resume documentación para facturar, no reemplaza los estados de la OC. Los importes se separan por moneda y fuente. La bandeja de expedientes por completar incluye pendientes anteriores al período y advierte esa diferencia respecto de los KPI del período.

Los conteos se obtienen desde SQL, con fuentes y permisos reales. Guía no necesita aceptación adicional, confirmar OC completa su trabajo y la preparación para Facturado admite OC completada sin ejecución separada. Las series de la curva se llaman Confirmación del cliente y Facturado. Cuentan confirmaciones/referencias vigentes según su fecha de registro, dentro del período consultado, aunque la OC fuera anterior. Se excluyen retiradas y OC canceladas. No se utiliza el primer evento histórico de cada OC; una confirmación reemplazada se cuenta por el registro vigente. Las correcciones conservan su fecha de registro.

## Archivos y trazabilidad

Adjuntos privados en Cloudflare R2, organizados por tipo y número: OCCliente/{número}, Guia/{número}, ConfirmacionCliente/{número}, ConfigHYT/Firma. Cada archivo conserva nombre, fecha y usuario; se visualiza o descarga desde su documento. La firma del presupuesto y los dos formatos HTML de impresión se mantienen configurables. Los cambios de estado y datos conservan auditoría interna y control de concurrencia, aunque el historial general no se muestre en las fichas.

## Usuarios, contraseñas y auditoría

En **Administración → Usuarios y permisos**, el administrador crea cuentas con nombres, apellidos, correo, perfil y contraseña temporal. Los perfiles agrupan permisos por módulo y acción. El primer ingreso y el ingreso tras un restablecimiento requieren cambiar la contraseña antes de entrar a los módulos. El acceso **Cambiar contraseña** está en la cabecera para todos los perfiles.

Los campos de **Cambiar contraseña** se muestran uno por fila y cada uno tiene su propio icono para mostrar u ocultar el contenido. El mismo icono se usa en el inicio de sesión y en el campo de contraseña temporal al registrar o restablecer un usuario.

**Administración → Auditoría** permite consultar cambios por fecha, usuario, módulo, acción y referencia. **Presupuestos** reúne los eventos del presupuesto y sus revisiones en un solo filtro. **Acción** muestra las opciones del módulo dentro del período seleccionado y conserva la elección al consultar.
