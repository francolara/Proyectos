namespace SistemaControlEspaciosDeportivosWeb.ViewModels.Ayuda;

// Firma: FRANCO LARA - 07/09/2026 | Centraliza y actualiza el catalogo FAQ de todos los modulos del panel operativo.
public static class AyudaCatalogoFactory
{
    public static AyudaIndexViewModel Crear(ModuloBaseViewModel baseVm, string? moduloSolicitado)
    {
        var categorias = ConstruirCategorias();
        var clave = NormalizarClave(moduloSolicitado);
        var seleccionado = categorias.SelectMany(x => x.Modulos)
            .FirstOrDefault(x => x.Clave.Equals(clave, StringComparison.OrdinalIgnoreCase))
            ?? categorias.SelectMany(x => x.Modulos).First(x => x.Clave == "DASHBOARD");
        var categoria = categorias.First(x => x.Modulos.Any(m => m.Clave == seleccionado.Clave));

        return new AyudaIndexViewModel
        {
            Base = baseVm,
            CategoriaSeleccionadaClave = categoria.Clave,
            ModuloSeleccionadoClave = seleccionado.Clave,
            ModuloSeleccionadoTitulo = seleccionado.Titulo,
            ModuloSolicitado = string.IsNullOrWhiteSpace(moduloSolicitado) ? null : moduloSolicitado.Trim(),
            TotalPreguntas = categorias.SelectMany(x => x.Modulos).Sum(x => x.Preguntas.Count),
            Categorias = categorias
        };
    }

    private static IReadOnlyCollection<AyudaCategoriaViewModel> ConstruirCategorias() =>
    [
        Categoria("GENERAL", "General", "bi-grid-1x2", "Control del negocio, accesos, configuración, plan y orientación del panel.",
        [
            Modulo("DASHBOARD", "Dashboard", "bi-speedometer2", "Indicadores y accesos rápidos para iniciar y controlar la jornada.",
                ("¿Qué reviso primero al iniciar el turno?", "Las reservas del día, pendientes de confirmación, ingresos, saldos por cobrar y alertas operativas."),
                ("¿Qué cambia al seleccionar una sede o fecha?", "Los indicadores, listados y resúmenes se recalculan para el contexto elegido; así evitas mezclar operaciones de sedes distintas."),
                ("¿Qué representa la salud operativa?", "Resume señales como ocupación, reservas pendientes, inasistencias y ritmo de cobro para ayudarte a priorizar."),
                ("¿Para qué sirven las acciones rápidas?", "Permiten iniciar tareas frecuentes, como crear una reserva, registrar un pago, emitir un comprobante o abrir reportes."),
                ("¿Puedo crear una reserva desde el Dashboard?", "Sí. El acceso de nueva reserva abre el mismo flujo operativo y respeta sede, espacio, horario, disponibilidad y permisos."),
                ("¿Cada cuánto conviene revisarlo?", "Al iniciar el turno, durante los cambios de mayor demanda y antes del cierre diario."),
                ("¿Qué alerta debo priorizar?", "Primero conflictos o reservas próximas sin confirmar; después saldos pendientes y tareas administrativas."),
                ("¿Cómo imprimo el resumen?", "Usa la opción de impresión del Dashboard para obtener una vista limpia del contexto actualmente filtrado.")),

            Modulo("USUARIOS", "Usuarios", "bi-people", "Equipo de trabajo, roles, sedes, recuperación de acceso y permisos por módulo.",
                ("¿Cómo incorporo a un trabajador?", "Asígnalo por correo, completa sus nombres, rol y sede. Si la cuenta no existe, el sistema inicia su registro según la configuración de acceso vigente."),
                ("¿Qué diferencia hay entre rol y permisos?", "El rol define el perfil general; los permisos determinan qué módulos puede ver y si puede crear, editar o eliminar."),
                ("¿Cómo limito a un usuario a una sede?", "Selecciona la sede asignada en su registro. El usuario trabajará dentro de ese alcance cuando el módulo aplique filtro por sede."),
                ("¿Puedo cambiar el rol después?", "Sí. Selecciona el nuevo rol y guarda el cambio desde la fila del usuario."),
                ("¿Por qué el botón Guardar aparece junto al rol y la sede?", "Porque confirma en una sola acción la asignación operativa mostrada en esa fila."),
                ("¿Cómo personalizo sus accesos?", "Abre Permisos, activa los módulos autorizados y define las acciones de consulta, creación, edición y eliminación."),
                ("¿Qué ocurre al desactivar a un usuario?", "Pierde el acceso operativo al negocio, pero se conserva su historial para auditoría."),
                ("¿Cómo envío acceso a alguien que aún no confirmó su correo?", "Usa Enviar enlace para generar el flujo de confirmación o recuperación disponible."),
                ("¿Debo compartir mi contraseña?", "No. Cada integrante debe usar su propia cuenta; así los permisos y la trazabilidad permanecen separados.")),

            Modulo("CONFIGURACION", "Configuración", "bi-sliders", "Identidad, datos fiscales y reglas que gobiernan reservas, pagos y emisión.",
                ("¿Qué debo configurar antes de operar?", "Nombre e identidad del negocio, sede, moneda, políticas de reserva y pago, datos fiscales y series de comprobantes."),
                ("¿Dónde cambio el logo y los datos comerciales?", "En Configuración del negocio. El logo y nombre se reutilizan en distintas vistas y documentos."),
                ("¿Para qué sirve el ubigeo fiscal?", "Identifica correctamente la ubicación tributaria usada en comprobantes y datos del emisor."),
                ("¿Qué controla el adelanto mínimo?", "Define el importe o porcentaje requerido por la política de confirmación antes de considerar asegurada una reserva."),
                ("¿Qué controla el máximo de horas para reserva pública?", "Limita con cuánta anticipación o dentro de qué ventana se aceptan solicitudes desde el portal público."),
                ("¿Cómo se maneja el IGV?", "La tasa configurada se usa en cálculos y comprobantes cuando corresponde. Revisa su valor antes de emitir."),
                ("¿Cómo habilito factura electrónica?", "Completa los datos fiscales y activa la emisión. También deben estar listos proveedor electrónico, series y credenciales requeridas."),
                ("¿Qué son las series de comprobantes?", "Son las numeraciones separadas para boletas, facturas y notas. Deben corresponder a la configuración tributaria vigente."),
                ("¿Qué hace la cancelación automática?", "Define cuándo una reserva pendiente podría liberarse por falta de confirmación. Solo actúa si el proceso programado correspondiente está habilitado en el alojamiento."),
                ("¿La zona horaria del hosting cambia mis horarios?", "No debería. La aplicación interpreta la operación en la zona del negocio y usa UTC para instantes técnicos cuando corresponde; el usuario sigue viendo la hora local."),
                ("¿Qué debo comprobar después de guardar?", "Vuelve a abrir la sección y verifica identidad, moneda, reglas de reserva, datos fiscales y series antes de operar.")),

            Modulo("MISUSCRIPCION", "Mi suscripción", "bi-credit-card-2-front", "Plan contratado, vigencia, límites y capacidad disponible.",
                ("¿Qué información muestra esta sección?", "Plan actual, estado, vigencia y límites operativos disponibles para el negocio."),
                ("¿Qué ocurre si alcanzo un límite?", "El sistema puede impedir nuevas altas del recurso afectado hasta ampliar el plan o liberar capacidad."),
                ("¿La suspensión elimina mis datos?", "No implica una eliminación automática; restringe el servicio según el estado comercial y las políticas de la plataforma."),
                ("¿Dónde reviso la fecha de renovación?", "En el detalle de vigencia de la suscripción."),
                ("¿Quién puede gestionar el plan?", "El administrador o titular autorizado del negocio."),
                ("¿Qué hago si el plan mostrado no coincide?", "No dupliques pagos ni registros. Comunícate con soporte indicando el negocio y el estado que observas.")),

            Modulo("AYUDA", "Ayuda", "bi-life-preserver", "FAQ contextual y buscador operativo para resolver dudas dentro del panel.",
                ("¿Cómo encuentro una respuesta?", "Selecciona una categoría y módulo o escribe una palabra en el buscador del módulo visible."),
                ("¿Qué significa ayuda contextual?", "Al abrir Ayuda desde un módulo, el catálogo se enfoca automáticamente en la sección relacionada."),
                ("¿La búsqueda revisa todo el sistema?", "Filtra las preguntas del módulo visible para mantener resultados precisos y fáciles de leer."),
                ("¿Cómo regreso a todas las preguntas?", "Borra el texto del buscador o cambia de módulo."),
                ("¿Qué hago si una respuesta no coincide con mi permiso?", "Tu rol o plan puede ocultar acciones. Consulta al administrador del negocio antes de asumir que existe un error."),
                ("¿Para qué sirve la Guía de configuración?", "Acompaña los pasos iniciales del negocio. Se muestra debajo de Notificaciones mientras haya tareas pendientes y desaparece al llegar al 100 %."))
        ]),

        Categoria("OPERACION", "Operación", "bi-calendar2-check", "Catálogos, infraestructura, agenda, cobros y documentos de la atención diaria.",
        [
            Modulo("MAESTROS", "Maestros", "bi-toggles", "Catálogos base utilizados por reservas, pagos y comprobantes.",
                ("¿Qué catálogos se administran aquí?", "Monedas, tipos de suelo, deportes, formas de pago, tipos de documento y series documentarias."),
                ("¿Por qué deben configurarse primero?", "Porque las sedes, espacios, clientes, cobros y comprobantes dependen de estos valores."),
                ("¿Puedo editar un maestro en uso?", "Sí cuando la pantalla lo permite, pero revisa su impacto histórico antes de cambiar significado o código."),
                ("¿Conviene eliminar un valor ya usado?", "No. Conserva la referencia histórica y evita retirar catálogos vinculados a operaciones existentes."),
                ("¿Para qué sirven los deportes y tipos de suelo?", "Clasifican los espacios y facilitan su presentación, búsqueda y operación."),
                ("¿Para qué sirven las formas de pago?", "Normalizan cómo se registran efectivo, transferencias, tarjetas u otros medios aceptados."),
                ("¿Qué diferencia hay entre tipo de documento y serie?", "El tipo identifica boleta, factura u otro documento; la serie controla su numeración."),
                ("¿Qué debo validar al terminar?", "Que cada catálogo necesario esté activo y que las series correspondan al tipo de comprobante correcto.")),

            Modulo("SEDES", "Sedes", "bi-geo-alt", "Locales del negocio, ubicación, contacto, horarios e imagen representativa.",
                ("¿Qué representa una sede?", "Un local físico desde el que se agrupan espacios, horarios, reservas y usuarios asignados."),
                ("¿Qué datos debo completar?", "Nombre, dirección, ubicación, contacto, horarios y demás información solicitada por el formulario."),
                ("¿La sede afecta lo que ve el usuario?", "Sí. Un usuario asignado puede operar dentro de esa sede y los filtros separan su información."),
                ("¿Puedo subir una imagen?", "Sí. La imagen ayuda a identificar y presentar el local en las vistas que la consumen."),
                ("¿Qué debo hacer antes de eliminarla?", "Revisa espacios, usuarios, reservas y operaciones vinculadas; una sede con dependencias puede no admitir eliminación."),
                ("¿Puedo editar ubicación y contacto?", "Sí. Los cambios se reflejan en las consultas posteriores del local."),
                ("¿Cómo organizo varias sedes?", "Usa nombres inequívocos y asigna correctamente espacios y personal a cada local."),
                ("¿Por qué una sede no aparece en una lista?", "Comprueba que pertenezca al negocio actual, esté disponible para tu usuario y no haya filtros activos.")),

            Modulo("ESPACIOS", "Espacios deportivos", "bi-building", "Canchas y ambientes con horarios, tarifas, imágenes, relaciones y reseñas.",
                ("¿Qué define a un espacio deportivo?", "Su sede, nombre, deporte, superficie, capacidad, horarios, precios, visibilidad y condiciones de uso."),
                ("¿Cómo se configuran los horarios?", "Cada espacio maneja sus tramos disponibles; la agenda usa esos horarios junto con bloqueos y reservas existentes."),
                ("¿Puedo usar tarifas distintas?", "Sí. El espacio admite su tarifa base y las variaciones habilitadas por día, horario o feriado según su configuración."),
                ("¿Qué es un espacio directo o compuesto?", "Permite relacionar ambientes que comparten físicamente una cancha completa o subdivisiones; reservar uno puede bloquear los relacionados."),
                ("¿Cómo evito cruces entre espacios relacionados?", "Configura correctamente la relación. La disponibilidad considera reservas y bloqueos del espacio y de sus componentes vinculados."),
                ("¿Cuántas imágenes puedo administrar?", "La galería admite hasta tres imágenes, con una principal para la presentación del espacio."),
                ("¿Qué diferencia hay entre activo y visible al público?", "Activo permite operarlo internamente; la visibilidad determina si se ofrece en el portal público."),
                ("¿Puedo cambiar la imagen principal?", "Sí. Selecciona la imagen que representará al espacio en listados y portal."),
                ("¿Cómo se gestionan las reseñas?", "Puedes revisarlas, responderlas u ocultarlas desde la administración, respetando la trazabilidad del comentario."),
                ("¿Puedo eliminar un espacio con historial?", "Verifica primero sus reservas, bloqueos y relaciones. En general conviene desactivarlo para conservar el historial.")),

            Modulo("RESERVAS", "Reservas", "bi-calendar-event", "Calendario, solicitudes web, cotización, disponibilidad, bloqueos y seguimiento.",
                ("¿Qué vistas ofrece Reservas?", "Calendario y listado con filtros por sede, espacio, fecha y estado, además de resúmenes operativos del día."),
                ("¿Cómo creo una reserva?", "Elige cliente, sede, espacio, fecha y horario; valida disponibilidad, precio y política de pago antes de guardar."),
                ("¿Puedo crear al cliente sin salir?", "Sí. El alta rápida permite registrar los datos esenciales y continuar con la reserva."),
                ("¿Cómo se evita una doble reserva?", "La validación revisa cruces de fecha y hora, bloqueos y relaciones entre espacios antes de confirmar."),
                ("¿Qué pasa con los espacios compuestos?", "La ocupación del espacio completo o de una parte bloquea los componentes relacionados según la configuración."),
                ("¿Puedo cotizar antes de guardar?", "Sí. La cotización calcula duración, tarifa, descuentos, cupón y total con los datos seleccionados."),
                ("¿Cómo se aplican promociones y cupones?", "Las promociones válidas pueden intervenir en el cálculo y el cupón se valida por fecha, uso, sede, espacio y reglas configuradas."),
                ("¿Puedo registrar un pago al crearla?", "Sí cuando el formulario lo habilita. El resumen diferencia total, pagos previos, pago actual y saldo."),
                ("¿Qué puedo modificar desde el calendario?", "Puedes abrir el detalle y, con permisos, mover o ajustar la reserva; el sistema vuelve a validar disponibilidad."),
                ("¿Cómo cambio rápidamente el estado?", "Usa la acción de estado disponible en calendario o tabla. Mantén la secuencia coherente con confirmación, atención, cancelación o inasistencia."),
                ("¿Qué muestra el historial?", "Los eventos y cambios registrados para entender la evolución de la reserva."),
                ("¿Para qué sirve la acción masiva?", "Permite aplicar la operación habilitada a varias reservas seleccionadas, por ejemplo comunicaciones o cambios controlados."),
                ("¿Qué son los bloqueos?", "Periodos sin venta por mantenimiento, evento interno u otra causa. Pueden crearse y eliminarse con permisos."),
                ("¿Cómo llegan las solicitudes del portal?", "El cliente envía una solicitud pública. El equipo puede aprobarla, rechazarla o convertirla en reserva según disponibilidad y política."),
                ("¿Las solicitudes ya ocupan definitivamente el horario?", "No necesariamente. Revisa su estado y confirma o convierte la solicitud para consolidar la operación."),
                ("¿Cómo se manejan las horas si el hosting está en otra zona?", "El negocio trabaja con su hora local; los instantes técnicos se normalizan para evitar que la zona del servidor desplace el calendario."),
                ("¿Qué reviso antes de cerrar el día?", "Reservas pendientes, inasistencias, saldos, bloqueos del día siguiente y solicitudes públicas sin atender.")),

            Modulo("PAGOS", "Pagos", "bi-cash-coin", "Cobros parciales o totales asociados a reservas y sus saldos.",
                ("¿Cómo registro un pago?", "Busca la reserva, revisa su total y saldo, selecciona forma de pago e ingresa importe, fecha y referencia cuando corresponda."),
                ("¿Se permiten pagos parciales?", "Sí. Puedes registrar varios abonos hasta completar el saldo de la reserva."),
                ("¿Puedo pagar más que el saldo?", "No debería permitirse. Verifica el resumen antes de guardar para no exceder el monto pendiente."),
                ("¿Para qué sirve el número de operación?", "Identifica transferencias, depósitos o transacciones electrónicas y facilita su conciliación."),
                ("¿Cómo encuentro un cobro?", "Usa los filtros de búsqueda, fechas, reserva o cliente disponibles en el listado."),
                ("¿Puedo corregir un pago?", "Sí con permiso de edición. Comprueba importe, medio, fecha y referencia antes de confirmar."),
                ("¿Qué pasa si elimino un pago?", "El saldo de la reserva se recalcula y el historial pierde ese cobro; hazlo solo si realmente fue registrado por error."),
                ("¿Registrar un pago emite comprobante?", "No necesariamente. El cobro y el comprobante son operaciones relacionadas pero independientes."),
                ("¿Cómo verifico el saldo real?", "Consulta el resumen de la reserva, que separa total, pagado y pendiente.")),

            Modulo("COMPROBANTES", "Comprobantes", "bi-receipt", "Boletas, facturas, notas, envío electrónico, impresión y entrega al cliente.",
                ("¿Qué necesito antes de emitir?", "Datos fiscales del negocio y cliente, serie, tipo de documento, detalle, importes y configuración del proveedor electrónico."),
                ("¿Cómo ubico la operación?", "Busca la reserva o cliente y abre su contexto antes de completar el comprobante."),
                ("¿Cuándo emito boleta o factura?", "Según el tipo de cliente y sustento tributario solicitado. Para factura valida RUC, razón social, dirección y ubigeo."),
                ("¿Puedo obtener una vista previa?", "Sí. Revisa datos, conceptos, impuestos y totales antes del envío definitivo."),
                ("¿Qué hace Enviar a SUNAT?", "Transmite el documento mediante el proveedor electrónico configurado y registra la respuesta disponible."),
                ("¿Qué hago si el envío falla?", "No dupliques el comprobante. Revisa el mensaje, conexión, credenciales, serie y datos tributarios antes de reintentar."),
                ("¿Para qué sirven las notas de crédito y débito?", "Corrigen o ajustan un comprobante ya emitido con el motivo y documento de referencia exigidos."),
                ("¿Puedo editar un comprobante enviado?", "Un documento tributario aceptado no se corrige como borrador; usa el procedimiento tributario correspondiente, normalmente una nota."),
                ("¿Cómo lo entrego al cliente?", "Puedes imprimir la representación, descargar los archivos disponibles o enviarlos por correo desde las acciones habilitadas."),
                ("¿Qué archivos electrónicos puedo descargar?", "Los que el proveedor haya generado y el sistema tenga disponibles, como representación impresa y archivos tributarios."),
                ("¿Registrar comprobante equivale a cobrar?", "No. Comprueba por separado que el pago de la reserva esté registrado."),
                ("¿Qué debo revisar al final del día?", "Documentos pendientes de envío, rechazados, notas emitidas y correlativos de las series."))
        ]),

        Categoria("COMERCIAL", "Comercial", "bi-megaphone", "Clientes, campañas, beneficios y análisis para impulsar ocupación e ingresos.",
        [
            Modulo("CLIENTES", "Clientes", "bi-person-vcard", "Datos de contacto, documento, ubicación, estado e historial de usuarios del servicio.",
                ("¿Qué datos conviene registrar?", "Documento, nombres o razón social, teléfono, correo y ubicación; completa datos fiscales si solicitará factura."),
                ("¿Cómo encuentro un cliente?", "Busca por nombre, documento, teléfono o correo y combina el filtro de estado cuando sea necesario."),
                ("¿Puedo registrar una empresa?", "Sí. Usa el tipo de documento y los campos fiscales apropiados para una persona jurídica."),
                ("¿Por qué es importante el correo?", "Se usa para comunicaciones, confirmaciones y envío de comprobantes cuando esos flujos están habilitados."),
                ("¿Puedo editar sus datos?", "Sí con permisos. Verifica especialmente documento, nombre fiscal, correo y ubigeo."),
                ("¿Qué pasa al desactivarlo?", "Se conserva su historial, pero puede dejar de estar disponible para operaciones nuevas según el flujo."),
                ("¿Conviene eliminar un cliente con reservas?", "No. Mantén la trazabilidad y usa el estado cuando exista historial relacionado."),
                ("¿Cómo evito duplicados?", "Busca primero por documento, correo y teléfono antes de crear un registro.")),

            Modulo("PROMOCIONES", "Promociones", "bi-percent", "Descuentos programados por vigencia, sede, espacio y franja horaria.",
                ("¿Qué define una promoción?", "Nombre, porcentaje de descuento, vigencia, horario y alcance por sede o espacio."),
                ("¿Puede aplicarse a una sola cancha?", "Sí. Selecciona la sede y el espacio cuando quieras restringir el beneficio."),
                ("¿Cómo se controla la vigencia?", "Con fechas de inicio y fin; fuera de ese periodo no debe intervenir en la cotización."),
                ("¿Puedo limitarla por horario?", "Sí, usando las horas configuradas en el formulario cuando corresponda."),
                ("¿Qué pasa si coinciden varias promociones?", "La cotización aplica las reglas vigentes del sistema; revisa el resultado antes de confirmar y evita campañas ambiguas."),
                ("¿Puedo editar una campaña activa?", "Sí con permisos, pero el cambio puede afectar nuevas cotizaciones. Conserva claridad sobre lo ya ofrecido."),
                ("¿Eliminarla cambia reservas existentes?", "No debería recalcular retroactivamente una reserva ya guardada; valida cada caso antes de retirar la campaña."),
                ("¿Cómo compruebo que funciona?", "Realiza una cotización dentro de la fecha, horario, sede y espacio definidos.")),

            Modulo("CUPONES", "Cupones", "bi-ticket-perforated", "Códigos promocionales con valor, vigencia, usos y alcance controlado.",
                ("¿Qué tipos de cupón existen?", "Porcentaje o monto fijo, según la opción seleccionada al crearlo."),
                ("¿Qué datos debo definir?", "Código, tipo, valor, máximo de usos, fechas y alcance por sede o espacio."),
                ("¿El código distingue mayúsculas?", "Para evitar confusión, comunica exactamente el código guardado y pruébalo antes de difundirlo."),
                ("¿Cómo se limita el número de usos?", "Con el máximo configurado; cada aplicación válida consume disponibilidad según las reglas del sistema."),
                ("¿Puede valer solo para una cancha?", "Sí. Restringe el cupón por sede y espacio cuando la campaña lo requiera."),
                ("¿Qué valida la reserva?", "Código, estado, vigencia, usos disponibles, alcance y compatibilidad con los datos cotizados."),
                ("¿Qué diferencia hay entre desactivar y eliminar?", "Desactivar detiene usos nuevos conservando el registro; eliminar debe reservarse para datos sin historial relevante."),
                ("¿Puedo editar un cupón difundido?", "Sí con cautela. Cambiar valor o condiciones puede generar diferencias con lo comunicado al cliente."),
                ("¿Cómo confirmo el descuento?", "Valida el cupón dentro de la reserva y revisa el desglose de la cotización antes de guardar.")),

            Modulo("REPORTES", "Reportes", "bi-bar-chart-line", "Lectura operativa y exportación de reservas, ingresos, ocupación y saldos.",
                ("¿Qué información puedo analizar?", "Los bloques disponibles resumen operación, ingresos, reservas, ocupación, clientes y saldos según el reporte seleccionado."),
                ("¿Los filtros afectan todos los bloques?", "Se aplican según el diseño de cada reporte. Confirma sede y periodo mostrados antes de interpretar resultados."),
                ("¿Cuál es la diferencia entre reservas e ingresos?", "Una reserva registra la operación comercial; el ingreso corresponde a pagos efectivamente registrados."),
                ("¿Cómo identifico saldos pendientes?", "Consulta el bloque de cobranza o reservas con diferencia entre total y monto pagado."),
                ("¿Puedo imprimir un reporte?", "Sí. La vista de impresión prepara el bloque y filtros seleccionados para una salida más limpia."),
                ("¿Puedo exportar los datos?", "Sí. Usa Exportar CSV para trabajar el bloque habilitado en una hoja de cálculo."),
                ("¿Por qué un pago reciente no aparece?", "Revisa el periodo, sede, estado, fecha usada por el reporte y actualiza la consulta."),
                ("¿Cómo interpreto la ocupación?", "Compara tiempo reservado con capacidad disponible en el periodo y contexto seleccionados."),
                ("¿Qué revisar antes de compartirlo?", "Sede, fechas, moneda, estados incluidos y si se trata de importes reservados, cobrados o pendientes."))
        ])
    ];

    private static AyudaCategoriaViewModel Categoria(string clave, string titulo, string icono, string descripcion, IReadOnlyCollection<AyudaModuloViewModel> modulos) =>
        new() { Clave = clave, Titulo = titulo, Icono = icono, Descripcion = descripcion, Modulos = modulos };

    private static AyudaModuloViewModel Modulo(string clave, string titulo, string icono, string resumen, params (string Pregunta, string Respuesta)[] preguntas) =>
        new()
        {
            Clave = clave,
            Titulo = titulo,
            Icono = icono,
            Resumen = resumen,
            Preguntas = preguntas.Select((x, indice) => new AyudaPreguntaViewModel
            {
                Id = $"faq-{clave.ToLowerInvariant()}-{indice + 1}",
                Pregunta = x.Pregunta,
                Respuesta = x.Respuesta
            }).ToArray()
        };

    private static string NormalizarClave(string? moduloSolicitado)
    {
        var clave = (moduloSolicitado ?? string.Empty).Trim().ToUpperInvariant();
        return clave switch
        {
            "PANEL" => "DASHBOARD",
            "ESPACIOSDEPORTIVOS" => "ESPACIOS",
            "SOLICITUDES" or "NOTIFICACIONES" => "RESERVAS",
            "CUENTA" => "USUARIOS",
            _ => clave
        };
    }
}
