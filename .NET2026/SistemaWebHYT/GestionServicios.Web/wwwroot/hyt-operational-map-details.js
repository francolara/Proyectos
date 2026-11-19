export const details = {
    operacion: {
        configuracion: {
            summary: 'Centraliza el catálogo de firmas y las series de venta que usan Presupuestos y Emisión de comprobantes.',
            rules: [
                'La consulta y edición requieren permisos de Configuración HYT.'
            ],
            validation: 'Si falta una firma o una serie de venta, abre la opción correspondiente de Configuración HYT.',
            result: 'Los parámetros quedan disponibles para el proceso que los utiliza.',
            modules: ['configuracion']
        },
        firmas: {
            summary: 'Administra el catálogo de imágenes de firma para las impresiones de Presupuesto.',
            rules: [
                'Las firmas pueden activarse o inactivarse. El catálogo admite una firma predeterminada.'
            ],
            validation: 'Si la firma no se imprime, revisa que esté activa y seleccionada en esa revisión. HYT omite las firmas inactivas.',
            result: 'Selecciona la firma en la revisión del Presupuesto y comprueba el resultado en la vista de impresión.',
            modules: ['firmas-presupuesto', 'presupuestos']
        },
        series: {
            summary: 'Mantiene las series disponibles para cada tipo de comprobante de venta.',
            rules: [
                'Las series de ventas se ofrecen según el tipo de comprobante. En compras se registra la serie del proveedor.'
            ],
            validation: 'Si una serie nueva no aparece en ventas, revisa el tipo elegido y las series configuradas. Los comprobantes históricos conservan su serie original.',
            result: 'La serie queda disponible en Emisión de comprobantes para el tipo configurado.',
            modules: ['series-venta', 'emision-comprobantes']
        },
        clientes: {
            summary: 'Identifica quién solicita el servicio y dónde se atenderá.',
            rules: [
                'El presupuesto se registra para un cliente y una tienda.',
                'Los contactos se consultan con el permiso de Clientes.'
            ],
            validation: 'Si la sede del servicio no aparece, revisa la tienda y su cliente antes de registrar el presupuesto.',
            result: 'Selecciona el cliente y la tienda en Presupuestos. Consulta los contactos cuando necesites identificar al interlocutor.',
            modules: ['clientes', 'tiendas', 'contactos', 'zonas']
        },
        catalogos: {
            summary: 'Los catálogos y las tarifas alimentan las partidas del presupuesto.',
            rules: [
                'Las tarifas con IGV se convierten a importes netos para el detalle.'
            ],
            validation: 'Si el precio sugerido difiere del esperado, revisa la unidad y la tarifa del material o servicio. Confirma si el precio de origen incluye IGV.',
            result: 'Usa las partidas del tarifario para preparar el detalle y revisar el importe neto del presupuesto.',
            modules: ['especialidades', 'tipos-servicio', 'unidades', 'materiales', 'tarifarios']
        },
        presupuesto: {
            summary: 'Punto de partida del expediente comercial.',
            rules: [
                'El correlativo visible es mensual. El tipo de atención es Invitación o Emergencia.',
                'El registro de una OC o Guía válidas puede aceptar el presupuesto automáticamente.'
            ],
            validation: 'Si el cliente observa el presupuesto, registra el motivo y corrige la propuesta. Antes de cancelar, revisa los documentos vinculados.',
            result: 'El presupuesto queda Registrado al guardar. Registra el envío desde el listado y vincula la OC del cliente o la Guía cuando corresponda.',
            modules: ['presupuestos']
        },
        orden: {
            summary: 'Registra la orden de compra y su referencia externa.',
            rules: [
                'Cada presupuesto admite como máximo una OC del cliente.',
                'El cierre operativo manual exige Guía y ejecución terminada.'
            ],
            validation: 'Si el expediente ya tiene una OC, revisa su ficha antes de intentar registrar otra. Para cerrar operativamente, comprueba la Guía y el término real de la ejecución.',
            result: 'La orden queda Registrada. Continúa con la documentación del trabajo y la confirmación del cliente.',
            modules: ['oc-cliente']
        },
        guia: {
            summary: 'Registra el documento de Guía y sus adjuntos.',
            rules: [
                'Cada presupuesto admite como máximo una Guía.',
                'En Emergencia se puede registrar la Guía sin OC previa.'
            ],
            validation: 'Si ya existe una Guía para el presupuesto, abre su ficha para revisar o corregir el número, la fecha y el PDF según tu permiso.',
            result: 'La Guía queda vinculada al presupuesto y disponible para consultar desde el expediente y sus reportes.',
            modules: ['guias']
        },
        ejecucion: {
            summary: 'Registra responsable, programación, inicio y término del trabajo.',
            rules: [
                'Confirmar la OC completa el trabajo de la orden. La ficha de ejecución se registra por separado.'
            ],
            validation: 'Distingue la fecha programada del inicio y término reales. Si el trabajo aún está en curso, actualiza el término cuando finalice.',
            result: 'El expediente conserva la información de ejecución para el seguimiento y el cierre operativo de la OC.',
            modules: ['presupuestos']
        },
        confirmacion: {
            summary: 'Acredita la conformidad del cliente con el trabajo.',
            rules: [
                'Solo puede existir una confirmación vigente por OC.',
                'Confirmar una OC registrada la lleva a Trabajo completado.'
            ],
            validation: 'Registra número y fecha de confirmación. Si el expediente tiene un marcado Facturado vigente, retira primero ese marcado antes de retirar la confirmación.',
            result: 'La confirmación queda visible en el expediente. Revisa la Guía y los requisitos de facturación.',
            modules: ['oc-cliente']
        },
        facturacion: {
            summary: 'Conecta el expediente con el proceso de ventas de Administración.',
            rules: [
                'La OC tiene una confirmación vigente del cliente.',
                'El presupuesto tiene una Guía registrada.',
                'La OC figura en Trabajo completado o Cerrada operativamente. Para expedientes históricos, HYT admite una ejecución terminada.'
            ],
            validation: 'Si HYT bloquea el marcado, registra la confirmación en la OC o la Guía en el expediente según el requisito pendiente. Para el trabajo completado, revisa la OC o la ejecución terminada en expedientes históricos.',
            result: 'El marcado conserva fecha, tipo, serie y número del comprobante externo. Corregirlo o retirarlo exige motivo. Continúa a Emisión de comprobantes para gestionar la venta vinculada.',
            modules: ['emision-comprobantes']
        },
        reportes: {
            summary: 'Consulta el expediente por Presupuesto u OC del cliente.',
            rules: [
                'Los datos relacionados se muestran según los permisos de consulta.',
                'Los totales de PEN y USD se presentan por separado.'
            ],
            validation: 'El estado disponible depende de la fuente seleccionada. Si faltan filas, revisa las fechas de registro, el texto de búsqueda y los pendientes activados, que se combinan entre sí.',
            result: 'Revisa las relaciones del expediente e imprime o exporta la consulta autorizada con los filtros aplicados.',
            modules: ['reportes']
        },
        ventas: {
            summary: 'Registra comprobantes de venta manualmente o mediante XML.',
            rules: [
                'Los borradores sin numeración quedan pendientes de facturar.',
                'Una nota de crédito manual referencia una factura del mismo cliente.',
                'Un comprobante de venta puede vincular varios presupuestos del mismo cliente y una OC vigente opcional.'
            ],
            validation: 'Al importar XML, HYT valida que su RUC figure como emisor y que la numeración no esté duplicada. Para una nota de crédito, comprueba la factura de referencia. Si falta el presupuesto o la OC en la selección, verifica el cliente y el estado del documento.',
            result: 'Una factura o boleta numerada y vinculada se refleja como facturada en el expediente. Una factura elegible puede pasar a Factoring.',
            modules: ['emision-comprobantes']
        },
        factoring: {
            summary: 'Agrupa facturas numeradas en lotes de negociación.',
            rules: [
                'Las facturas de un lote tienen la misma moneda y cada factura pertenece a un único lote.',
                'La fecha de desembolso por sí sola no acredita el pago.'
            ],
            validation: 'Si una factura no está disponible, verifica que esté numerada, que coincida con la moneda y que aún no pertenezca a otro lote.',
            result: 'La venta muestra su vínculo con el lote. Consulta allí la negociación y la fecha de desembolso si se registró.',
            modules: ['factoring']
        },
        controles: {
            summary: 'Administra usuarios, perfiles y trazabilidad.',
            rules: [
                'Los perfiles agrupan permisos por módulo y acción.',
                'Auditoría permite filtrar eventos por usuario, fecha, módulo y acción.'
            ],
            validation: 'Valida que el usuario tenga el perfil apropiado. Las cuentas con contraseña temporal deben cambiarla antes de entrar a los módulos.',
            result: 'Usa Usuarios y permisos para administrar los accesos y Auditoría para revisar los cambios.',
            modules: ['usuarios', 'auditoria']
        }
    },
    planillas: {
        ceses: {
            summary: 'Desde la ficha de Trabajadores se registra el cese y se consulta el historial de ceses y reingresos.',
            rules: [
                'El registro de cese o reingreso requiere el permiso correspondiente.'
            ],
            validation: 'Para reingresar, usa una fecha posterior al cese vigente y que no sea posterior a hoy en Lima.',
            result: 'El reingreso reactiva la ficha y actualiza su fecha de ingreso, con trazabilidad en el historial.',
            modules: ['trabajadores']
        },
        telecredito: {
            summary: 'Genera el TXT de abonos de Telecrédito BCP desde la Planificación de Planilla.',
            rules: [
                'Requiere un plan guardado y pendiente, permisos de pagos y acceso a datos de cuentas. Solo se exportan las filas incluidas.'
            ],
            validation: 'Revisa la cuenta de cargo de HYT y los datos de abono del trabajador. Para bancos distintos de BCP se utiliza el CCI.',
            result: 'Descarga el TXT para procesar los abonos. Registra la confirmación en HYT cuando el pago se haga efectivo.',
            modules: ['planificacion-pagos', 'telecredito-config']
        },
        maestros: {
            summary: 'Cargos y conceptos estructuran las fichas y movimientos del personal.',
            rules: [
                'El sueldo base de la ficha sirve como referencia para planificar el pago.'
            ],
            validation: 'Revisa que el cargo esté activo antes de asignarlo. Una ficha anterior puede conservar el cargo inactivo que ya tenía.',
            result: 'Asigna el cargo y revisa el sueldo de referencia en Trabajadores. Consulta los conceptos desde su catálogo.',
            modules: ['cargos', 'conceptos-planilla']
        },
        trabajadores: {
            summary: 'Mantiene la ficha del trabajador, su vínculo y frecuencia de pago.',
            rules: [
                'Los datos bancarios se muestran a perfiles con permiso específico.'
            ],
            validation: 'Si alguien no aparece al planificar, revisa si está activo, su fecha de ingreso y la frecuencia de pago de la ficha.',
            result: 'Con la ficha al día, consulta los procesos de asistencia, vacaciones o planificación que correspondan.',
            modules: ['trabajadores']
        },
        asistencia: {
            summary: 'Gestiona la asistencia y el tiempo extra del trabajador.',
            rules: [
                'La planificación toma los minutos extra aprobados del período.'
            ],
            validation: 'Si las horas extra no aparecen en el plan, comprueba que estén aprobadas y que su fecha esté dentro del período elegido.',
            result: 'Los minutos extra aprobados quedan disponibles como referencia en la planificación del período.',
            modules: ['trabajadores']
        },
        vacaciones: {
            summary: 'Desde la ficha de Trabajadores se registran las entradas de días reconocidos y las salidas de vacaciones del personal en planilla.',
            rules: [
                'Las salidas aprobadas, gozadas o programadas descuentan saldo.'
            ],
            validation: 'Revisa los días reconocidos y las salidas registradas al comprobar el saldo. Una salida pendiente de aprobación o anulada no lo descuenta.',
            result: 'Consulta el saldo en Vacaciones desde la ficha del trabajador antes de programar una salida.',
            modules: ['trabajadores']
        },
        sctr: {
            summary: 'Registra pólizas de salud y pensión y sus trabajadores asignados.',
            rules: [
                'Una póliza puede cubrir a varios trabajadores. La vigencia se evalúa con la fecha de Lima.'
            ],
            validation: 'Si el trabajador no aparece cubierto, revisa su asignación y las fechas. HYT distingue pólizas programadas, vigentes, próximas a vencer y vencidas.',
            result: 'La ficha conserva las asignaciones. Consulta el listado de SCTR para revisar la cobertura y los vencimientos.',
            modules: ['sctr']
        },
        plan: {
            summary: 'Prepara importes por semana, quincena o mes.',
            rules: [
                'Precarga personal activo con la frecuencia seleccionada. El neto considera base, extras y descuentos.',
                'Guardar el plan deja el período listo para revisión.',
                'Telecrédito BCP exige un plan guardado y pendiente, permiso de consulta de cuentas y permisos de pagos. El TXT incluye solo las filas seleccionadas.'
            ],
            validation: 'Antes de exportar o confirmar, guarda los ajustes y las incorporaciones. Revisa los trabajadores incluidos y, para Telecrédito, los datos de abono autorizados.',
            result: 'El plan guardado queda pendiente para revisión. Cuando el pago se haga efectivo, regístralo en Confirmación de pago.',
            modules: ['planificacion-pagos']
        },
        pago: {
            summary: 'Confirma un plan guardado con fecha y número de operación.',
            rules: [
                'La confirmación registra los movimientos de personal y cierra el período en una transacción.',
                'Un período pagado se conserva para consulta y reportes.'
            ],
            validation: 'Si HYT detecta movimientos superpuestos, revisa el estado de cuenta y la cobertura del período antes de reintentar. Comprueba la fecha y la referencia del pago.',
            result: 'El período queda pagado y bloqueado para edición. Consulta sus movimientos y el reporte mensual.',
            modules: ['planificacion-pagos']
        },
        obligaciones: {
            summary: 'Controla las obligaciones mensuales de AFP y ONP.',
            rules: [
                'El estimado es una proyección. El importe definitivo puede ajustarse antes de registrar el pago.'
            ],
            validation: 'Revisa el mes, el régimen y la entidad previsional. Contrasta el estimado con el importe definitivo antes de registrar fecha y referencia de pago.',
            result: 'Consulta la obligación por mes y estado en Reportes de Planillas.',
            modules: ['obligaciones-planillas']
        },
        reportes: {
            summary: 'Consulta Planillas, Obligaciones y Pagos a Proveedores según el tipo seleccionado.',
            rules: [
                'La consulta aplica permisos específicos de cada fuente. El cierre mensual requiere su propio permiso.'
            ],
            validation: 'Elige Planillas, Obligaciones o Pagos a Proveedores y verifica sus filtros. El cierre mensual requiere que el personal del mes esté terminado y un permiso específico.',
            result: 'Revisa la consulta y usa la impresión o exportación disponible para los datos que tu perfil puede consultar.',
            modules: ['reporte-planillas']
        },
        configuracion: {
            summary: 'Define la cuenta corriente de cargo y su moneda para el archivo de abonos de Telecrédito BCP.',
            rules: [
                'La cuenta de cargo BCP tiene 13 dígitos. La moneda se valida contra la cuenta.'
            ],
            validation: 'Si la exportación no puede usar la cuenta de cargo, revisa el número y la moneda en Configuración HYT.',
            result: 'La cuenta configurada queda disponible en Exportar Telecrédito, desde un período guardado y pendiente.',
            modules: ['telecredito-config']
        },
        proveedor: {
            summary: 'Mantiene los datos del tercero que realiza el trabajo o suministra bienes.',
            rules: [
                'La consulta y edición de datos bancarios requiere el permiso correspondiente.'
            ],
            validation: 'Busca primero por RUC o nombre. Si necesitas consultar o editar los datos bancarios, revisa que el perfil tenga el permiso correspondiente.',
            result: 'La ficha queda disponible para seleccionar al proveedor en trabajos y registros de compras.',
            modules: ['proveedores']
        },
        trabajos: {
            summary: 'Registra el costo del proveedor asociado al servicio.',
            rules: [
                'Para vincularlo a una compra debe coincidir el proveedor y la moneda.'
            ],
            validation: 'Si el trabajo no aparece al vincularlo, comprueba el proveedor, la moneda y si el costo ya está asociado a otra compra.',
            result: 'El trabajo documenta el costo del servicio. Si corresponde, vincúlalo desde Registro de compras.',
            modules: ['pagos-proveedores']
        },
        compras: {
            summary: 'Registra compras y vincula sus trabajos de proveedor.',
            rules: [
                'Cada costo pertenece a una sola provisión y el total vinculado tiene como límite el importe del comprobante.'
            ],
            validation: 'En un XML de compra, HYT exige que el RUC de HYT sea el receptor. Si falla el vínculo, revisa proveedor, moneda y total de los costos seleccionados.',
            result: 'Con la compra registrada, consulta el saldo del proveedor y revisa si corresponde un depósito de detracción.',
            modules: ['provisiones-compras']
        },
        pagos: {
            summary: 'Consulta los saldos y registra los pagos que corresponden al proveedor.',
            rules: [
                'El saldo de una provisión descuenta el equivalente de la detracción en la moneda del documento.'
            ],
            validation: 'Antes de registrar el pago, revisa el saldo y si el trabajo está vinculado a una compra. Para una provisión, registra el pago desde su comprobante.',
            result: 'El registro conserva fecha, importe y referencia. Consulta el saldo actualizado o revisa el reporte de proveedores.',
            modules: ['pagos-pendientes-proveedores']
        },
        'reportes-proveedores': {
            summary: 'Consulta los pagos a proveedores registrados en HYT por mes, proveedor o presupuesto.',
            rules: [
                'La consulta presenta los pagos registrados y el saldo al cierre.',
                'La consulta requiere los permisos de reportes y del módulo de proveedores.'
            ],
            validation: 'Si no aparece la consulta de proveedores, revisa los permisos del perfil. Si aparece sin filas, verifica el mes, el proveedor y el filtro de saldo.',
            result: 'En Reportes de Planillas, elige Pagos a Proveedores para consultar, imprimir o exportar.',
            modules: ['reporte-planillas']
        },
        detracciones: {
            summary: 'Consulta pendientes y registra depósitos de detracción.',
            rules: [
                'El depósito conserva fecha, importe y número de operación. El saldo determina si queda pagada.'
            ],
            validation: 'Compara el importe y la moneda del depósito con los datos del comprobante. Si queda saldo, revisa los depósitos ya registrados y sus referencias.',
            result: 'La consulta identifica detracciones pendientes y pagadas, con la fecha y la referencia registradas.',
            modules: ['pago-detracciones']
        }
    }
};
