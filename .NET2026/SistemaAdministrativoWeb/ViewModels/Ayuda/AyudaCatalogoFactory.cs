namespace SistemaAdministrativoWeb.ViewModels.Ayuda;

// Firma: FRANCO LARA - 31/07/2026 | Completa la ayuda operativa General con los modulos Usuarios y Configuracion y sus preguntas contextuales.
// Firma: FRANCO LARA - 04/08/2026 | Actualiza la ayuda de Libros Electronicos con validacion integrada, archivos complementarios, periodos vacios y control de presentacion.
// Firma: FRANCO LARA - 05/08/2026 | Actualiza la ayuda del dashboard para incluir la tendencia historica de movimientos bancarios.
// Firma: FRANCO LARA - 25/08/2026 | Alinea la ayuda del cierre anual y de las cargas contables por defecto con el comportamiento vigente de cada modulo.
// Firma: FRANCO LARA - 26/08/2026 | Completa la ayuda operativa y documenta la seleccion emergente de rangos contables de cualquier nivel en Analisis, Libro Diario, Libro Mayor y Balance.
public static class AyudaCatalogoFactory
{
    public static AyudaIndexViewModel Crear(string? moduloSolicitado)
    {
        var categorias = ConstruirCategorias();
        var moduloNormalizado = NormalizarClave(moduloSolicitado);
        var moduloSeleccionado = categorias
            .SelectMany(x => x.Modulos)
            .FirstOrDefault(x => string.Equals(x.Clave, moduloNormalizado, StringComparison.OrdinalIgnoreCase))
            ?? categorias.SelectMany(x => x.Modulos).First(x => x.Clave == "DASHBOARD");

        var categoriaSeleccionada = categorias.First(x => x.Modulos.Any(m => m.Clave == moduloSeleccionado.Clave));

        return new AyudaIndexViewModel
        {
            CategoriaSeleccionadaClave = categoriaSeleccionada.Clave,
            ModuloSeleccionadoClave = moduloSeleccionado.Clave,
            ModuloSeleccionadoTitulo = moduloSeleccionado.Titulo,
            ModuloSolicitado = string.IsNullOrWhiteSpace(moduloSolicitado) ? null : moduloSolicitado.Trim(),
            TotalPreguntas = categorias.SelectMany(x => x.Modulos).Sum(x => x.Preguntas.Count),
            Categorias = categorias
        };
    }

    private static IReadOnlyCollection<AyudaCategoriaViewModel> ConstruirCategorias()
    {
        return
        [
            CrearCategoria(
                "GENERAL",
                "General",
                "bi-grid-1x2",
                "Orientacion del panel, empresas activas, soporte y estado de suscripcion.",
                [
                    CrearModulo("DASHBOARD", "Dashboard", "bi-speedometer2", "Resumen ejecutivo del periodo con KPIs, graficos y distribucion operativa.", CrearFaqDashboard()),
                    CrearModulo("EMPRESAS", "Empresas", "bi-buildings", "Seleccion y alta de empresas de trabajo dentro de la cuenta administradora.", CrearFaqEmpresas()),
                    CrearModulo("USUARIOS", "Usuarios", "bi-people", "Alta, vinculacion, empresas asignadas, roles y permisos de acceso por usuario.", CrearFaqUsuarios()),
                    CrearModulo("CONFIGURACION", "Configuracion", "bi-gear", "Datos de facturacion y preferencias de comprobante de la cuenta administradora.", CrearFaqConfiguracion()),
                    CrearModulo("MISUSCRIPCION", "Mi suscripcion", "bi-credit-card-2-front", "Estado de plan, vigencia, limites y datos de la cuenta administradora.", CrearFaqMiSuscripcion()),
                    CrearModulo("AYUDA", "Ayuda", "bi-life-preserver", "FAQ operativo con acceso contextual desde cada modulo.", CrearFaqAyuda())
                ]),
            CrearCategoria(
                "MANTENIMIENTO",
                "Mantenimiento",
                "bi-gear",
                "Catalogos contables, configuraciones base y parametros previos a la operacion diaria.",
                [
                    CrearModulo("PLANCUENTA", "Plan de cuentas", "bi-diagram-3", "Catalogo jerarquico de cuentas contables por empresa.", CrearFaqMantenimiento(
                        "Plan de cuentas",
                        "definir la estructura contable y sus niveles para compras, ventas, bancos, ajustes y reportes",
                        "la empresa activa, el nivel o grado objetivo y las reglas de comportamiento contable",
                        "codigo de cuenta, descripcion, grado, cuenta padre, moneda fija, columna de balance y checks de analisis o movimiento",
                        "cuando necesitas una cuenta nueva para una operacion, una reclasificacion o una expansion del plan actual",
                        "un codigo duplicado, un grado inconsistente o marcar una cuenta de resumen como cuenta de movimiento",
                        "las provisiones, los asientos manuales, las reglas de destino, los procesos de cierre y todos los reportes del mayor y balance",
                        "consultar por grado, revisar jerarquia y confirmar que la cuenta aparece en busquedas operativas y reportes",
                        "el responsable contable o quien gobierna la estructura del plan de cuentas",
                        "asientos, compras, ventas, caja y bancos, balance de comprobacion, libro diario y libro mayor",
                        ("Que carga el boton Cargar configuracion contable por defecto", "Carga en una sola transaccion el plan base, parametros de empresa, reglas de cuentas destino, cuentas de impuestos y cuentas por documento."),
                        ("La configuracion contable por defecto se carga al crear una empresa", "No. La empresa se registra sin disparar esta inicializacion; la carga se ejecuta expresamente desde Plan de cuentas cuando la empresa todavia no tiene un plan."),
                        ("De donde salen las cuentas de la carga por defecto", "Salen de las tablas maestras internas. Las configuraciones maestras guardan codigos contables y el proceso los resuelve contra las cuentas que acaba de crear para la empresa."),
                        ("Que ocurre si falta una cuenta requerida por una configuracion maestra", "La carga se detiene, muestra el codigo faltante y revierte toda la transaccion para no dejar una configuracion parcial."),
                        ("Puedo ejecutar otra vez la carga por defecto si la empresa ya tiene plan", "No. El proceso protege el plan existente y muestra que la empresa ya tiene cuentas registradas; cualquier ajuste posterior debe realizarse desde los mantenimientos correspondientes."))),
                    CrearModulo("CENTROCOSTO", "Centros de costo", "bi-diagram-2", "Clasificacion analitica para distribuir gastos y movimientos.", CrearFaqMantenimiento(
                        "Centros de costo",
                        "segmentar gastos e ingresos por area, sede, proyecto o unidad de negocio",
                        "la estructura analitica aprobada por la empresa y las cuentas que requieren centro de costo",
                        "codigo, descripcion, estado y criterio de uso en cuentas y movimientos",
                        "cuando la gerencia necesita separar resultados por unidad o controlar gastos especificos",
                        "crear centros sin una nomenclatura clara o seguir usando uno desactivado en operaciones nuevas",
                        "los asientos analiticos, las compras, los reportes por cuenta y los procesos de ajuste de cuentas",
                        "registrar un movimiento de prueba y verificar que el centro quede visible en consulta y reportes",
                        "contabilidad con coordinacion de administracion o control de gestion",
                        "asientos, compras, procesos de ajuste y analisis de cuentas")),
                    CrearModulo("CUENTACORRIENTE", "Cuentas corrientes", "bi-bank", "Bancos, cajas y cuentas operativas para flujos monetarios.", CrearFaqMantenimiento(
                        "Cuentas corrientes",
                        "definir las cuentas bancarias y cajas desde donde se registran ingresos, egresos y transferencias",
                        "la empresa activa, la moneda operativa y el vinculo con banco, caja u operacion financiera",
                        "numero de cuenta, banco, moneda, saldo inicial, estado y cuenta contable relacionada",
                        "cuando se incorpora una nueva cuenta bancaria, caja chica o fondo operativo",
                        "omitir la moneda correcta, no vincular la cuenta contable o dejar una cuenta cerrada como activa",
                        "caja y bancos, transferencias entre cuentas, conciliacion de saldos y reportes de movimientos bancarios",
                        "consultar caja y bancos con esa cuenta y confirmar que habilita registros y saldos correctos",
                        "tesoreria o contabilidad encargada del flujo bancario",
                        "caja y bancos, transferencias, asientos y dashboard bancario")),
                    CrearModulo("PERSONA", "Personas", "bi-people", "Clientes, proveedores y terceros vinculados a documentos y auxiliares.", CrearFaqMantenimiento(
                        "Personas",
                        "mantener clientes, proveedores y terceros para compras, ventas, aplicaciones y movimientos auxiliares",
                        "tipo de persona, documento de identidad o RUC y la condicion comercial del tercero",
                        "tipo de persona, documento, razon social o nombre, direccion, correo, telefono y estado",
                        "cuando aparece un nuevo proveedor, cliente o tercero que participara en un documento o asiento",
                        "duplicar personas por variaciones menores del nombre o registrar documentos incompletos",
                        "compras, ventas, aplicaciones, caja y bancos y reportes por auxiliar o documento",
                        "buscar por documento y verificar que la persona quede disponible en registros y consultas",
                        "administracion comercial o contabilidad segun el origen del tercero",
                        "compras, ventas, aplicaciones, analisis de cuentas y libro mayor auxiliar")),
                    CrearModulo("TIPOCAMBIO", "Tipos de cambio", "bi-currency-exchange", "TC por fecha y moneda para operaciones multimoneda.", CrearFaqMantenimiento(
                        "Tipos de cambio",
                        "registrar y sincronizar el TC usado por compras, ventas, bancos y reportes en moneda extranjera",
                        "la cuenta administradora, la fecha, la moneda y la politica de sincronizacion con TC oficial",
                        "fecha, moneda, compra, venta, promedio o TC operativo segun la implementacion vigente",
                        "cuando falta el TC del dia, se corrige un valor o se prepara un periodo con operaciones en USD",
                        "usar una fecha equivocada, repetir TC del mismo dia o no sincronizar antes de registrar documentos",
                        "compras, ventas, transferencias en distintas monedas, diferencia en cambio y reportes valorizados",
                        "consultar el periodo y comprobar que el TC aparece para la fecha requerida por la operacion",
                        "contabilidad o tesoreria con control diario del mercado cambiario",
                        "compras, ventas, caja y bancos, diferencia en cambio y libros electronicos")),
                    CrearModulo("ORIGEN", "Origenes", "bi-journal-richtext", "Subdiarios y origenes de asiento usados por cada circuito.", CrearFaqMantenimiento(
                        "Origenes",
                        "definir subdiarios y codigos de origen para asientos manuales y automaticos",
                        "el modulo que usara el origen y si debe permitir registros manuales",
                        "codigo de origen, nombre, modulo de origen, permiso de registro manual y estado",
                        "cuando abres un nuevo circuito operativo o separas documentos por subdiario",
                        "mezclar escenarios distintos bajo un mismo origen o dejar un origen inactivo en configuraciones",
                        "asientos manuales, compras, ventas, procesos mensuales y reportes por origen",
                        "registrar un asiento o documento de prueba y revisar que el origen se pueda seleccionar",
                        "contabilidad encargada del libro diario y de la trazabilidad por subdiario",
                        "asientos, compras, ventas, procesos contables y libro diario por origen",
                        ("Que hace Cargar origenes por defecto", "Carga los subdiarios desde el maestro de origenes y, en la misma transaccion, crea la configuracion contable inicial de la empresa para los escenarios maestros activos."),
                        ("Como relaciona la carga un origen maestro con la empresa", "La configuracion maestra guarda CodigoOrigen y el proceso busca ese mismo codigo entre los origenes creados para obtener el IdOrigen propio de la empresa."),
                        ("Puedo cargar origenes por defecto si la empresa ya tiene origenes", "No. Para evitar mezclar catalogos, el proceso se detiene cuando encuentra cualquier origen registrado para la empresa."),
                        ("Que ocurre si una configuracion maestra usa un origen inexistente", "La carga informa el CodigoOrigen faltante y revierte tanto los origenes como las configuraciones creadas durante el intento."))),
                    CrearModulo("CUENTADESTINOREGLA", "Cuentas destino", "bi-sliders2", "Reglas de contrapartida y distribucion automatica.", CrearFaqMantenimiento(
                        "Cuentas destino",
                        "asignar cuentas de contrapartida o destino para procesos automaticos y distribuciones contables",
                        "las reglas del negocio, la cuenta origen y el escenario donde se aplicara la derivacion",
                        "cuenta origen, cuentas destino de cargo y abono, orden, porcentaje, estado y observacion",
                        "cuando una operacion necesita generar automaticamente su contrapartida o reclasificacion",
                        "guardar porcentajes que no suman 100 o usar cuentas inactivas, de resumen o ajenas al plan empresarial",
                        "compras, asientos manuales, caja y bancos, ajuste de cuentas y diferencia en cambio cuando aplican distribuciones contables",
                        "ejecutar el proceso relacionado y revisar que la cuenta contrapartida resultante sea la esperada",
                        "contabilidad funcional o quien administra automatismos contables",
                        "compras, asientos, caja y bancos, ajuste de cuentas, diferencia en cambio y configuracion contable",
                        ("Las cuentas destino se configuran por ejercicio", "No. Existe una sola regla por empresa y cuenta origen; la misma configuracion se reutiliza mientras permanezca activa."),
                        ("Que cuentas puedo usar en una regla de destino", "Solo cuentas activas y de movimiento pertenecientes al plan contable de la empresa, tanto para el origen como para el cargo y el abono."),
                        ("Cuanto deben sumar los porcentajes del detalle", "Los detalles activos deben sumar exactamente 100 para que la distribucion contable quede completa."),
                        ("El Asiento de cierre usa cuentas destino", "No. El cierre vigente invierte directamente las cuentas de Inventario y no agrega cuentas destino, contrapartidas ni lineas de cuadre."))),
                    CrearModulo("CONFIGURACIONCONTABILIZACION", "Configuracion contable", "bi-gear-wide-connected", "Parametros de provision, impuestos, documentos y escenarios contables.", CrearFaqMantenimiento(
                        "Configuracion contable",
                        "gobernar como se contabilizan compras, ventas y procesos automaticos mediante configuraciones propias de cada empresa",
                        "los origenes operativos, cuentas por documento, impuestos, parametros contables y escenarios de proceso",
                        "origen por modulo y escenario, cuentas por documento, cuentas de impuesto, parametros contables y estados",
                        "cuando habilitas o corriges un escenario de compra, venta, aplicacion, ajuste, apertura, cierre u otro proceso automatico",
                        "cambiar cuentas sin validar el circuito completo o mezclar cuentas de impuestos y documentos",
                        "compras, ventas, aplicaciones, asientos automaticos, validacion tributaria y procesos contables",
                        "ejecutar el escenario configurado y revisar el origen y las cuentas que genera el asiento",
                        "contabilidad funcional con conocimiento del circuito tributario y contable",
                        "compras, ventas, aplicaciones, procesos, asientos automaticos, reportes y libros electronicos",
                        ("Las operaciones usan directamente las tablas maestras", "No. Compras, ventas y procesos consultan exclusivamente la configuracion de la empresa; los maestros solo sirven para las acciones expresas de carga inicial."),
                        ("Que ocurre si una operacion no tiene su cuenta empresarial configurada", "La operacion debe detenerse y mostrar que falta configurar la cuenta correspondiente; no debe completar el asiento buscando una alternativa en el maestro."),
                        ("Para que sirve la configuracion CIE", "Define el origen empresarial que utiliza el Asiento de cierre para generar su asiento automatico anual.")))
                ]),
            CrearCategoria(
                "REGISTRO",
                "Registro",
                "bi-journal-check",
                "Operaciones diarias que generan movimiento contable y afectan saldos, auxiliares o bancos.",
                [
                    CrearModulo("ASIENTO", "Asientos", "bi-receipt", "Registro manual de asientos del libro diario.", CrearFaqRegistro(
                        "Asientos",
                        "registrar movimientos manuales de diario cuando no vienen de un circuito automatico",
                        "el origen contable, fecha, glosa y el detalle de cuentas, debe y haber",
                        "origen, glosa, cuenta, centro de costo, tipo de documento, referencia, auxiliar y montos",
                        "el asiento puede guardarse aunque Debe y Haber sean diferentes; la diferencia visible debe revisarse y quedar sustentada por el responsable contable",
                        "crear un asiento manual innecesario para algo que ya genera compra, venta o proceso automatico",
                        "la consulta del libro diario, el libro mayor, balance de comprobacion y libros electronicos",
                        "revisar la consulta del periodo para confirmar correlativo, origen y montos",
                        "contabilidad responsable del libro diario y ajustes manuales",
                        "libro diario, libro mayor, balance, cierre y libros electronicos",
                        "asiento contable manual",
                        "el detalle del asiento",
                        "libro diario",
                        "debe y haber",
                        "origen y glosa",
                        "el periodo abierto")),
                    CrearModulo("COMPRA", "Compras", "bi-cart-check", "Provision de compras, saldo por pagar y validacion CPE.", CrearFaqRegistro(
                        "Compras",
                        "registrar comprobantes de proveedores y su provision contable",
                        "periodo abierto, configuracion contable del escenario, proveedor, tipo de documento y TC si aplica",
                        "proveedor, fecha, documento, moneda, escenario, detalle IGV, importe y saldo",
                        "la compra impacta cuentas por pagar, gasto o activo, impuestos y el control de saldo del proveedor",
                        "omitir escenario contable, registrar un documento repetido o usar TC inconsistente",
                        "registro de compras, libro diario, libro mayor, balance y aplicaciones de notas de credito",
                        "consultar la compra guardada, revisar saldo, estado CPE y asiento asociado si existe",
                        "cuentas por pagar o contabilidad de compras",
                        "registro de compras, aplicaciones, libro mayor, balance y dashboard",
                        "compra o provision",
                        "el detalle del comprobante",
                        "registro de compras",
                        "saldo por pagar",
                        "proveedor y documento",
                        "el periodo abierto")),
                    CrearModulo("VENTA", "Ventas", "bi-cash-stack", "Registro de ventas y asiento automatico de ingresos.", CrearFaqRegistro(
                        "Ventas",
                        "registrar comprobantes emitidos a clientes y su impacto en ingresos e impuestos",
                        "periodo abierto, cliente activo, configuracion contable por escenario y serie o documento correcto",
                        "cliente, fecha, comprobante, moneda, detalle, impuesto, total y saldo pendiente",
                        "la venta afecta cuentas por cobrar, ingresos, impuestos y reportes de ventas del periodo",
                        "usar un cliente incorrecto, duplicar serie y numero o no validar la configuracion del escenario",
                        "registro de ventas, libro diario, libro mayor, balance y analisis por cliente",
                        "consultar la venta emitida, validar saldo, asiento y lectura en reportes",
                        "facturacion o contabilidad de ingresos",
                        "registro de ventas, libro mayor, analisis de cuentas y dashboard",
                        "venta o comprobante emitido",
                        "el detalle del comprobante",
                        "registro de ventas",
                        "saldo por cobrar",
                        "cliente y comprobante",
                        "el periodo abierto")),
                    CrearModulo("CAJABANCO", "Caja y Bancos", "bi-cash-coin", "Movimientos de cuentas corrientes y flujo de tesoreria.", CrearFaqRegistro(
                        "Caja y Bancos",
                        "registrar ingresos, egresos y operaciones bancarias sobre una cuenta corriente",
                        "cuenta corriente activa, periodo abierto, tipo de operacion, fecha y tipo de cambio si corresponde",
                        "cuenta corriente, operacion, persona vinculada, numero de operacion, glosa, importe y detalle",
                        "el movimiento actualiza flujo bancario, saldos y puede generar el asiento asociado del circuito",
                        "dejar una cuenta corriente sin seleccionar, usar una operacion incompatible o cuadrar mal el detalle",
                        "consulta de caja y bancos, saldos por cuenta, transferencias y reportes auxiliares",
                        "consultar por cuenta y periodo para validar que el movimiento y los saldos se reflejan",
                        "tesoreria o caja con supervision contable",
                        "caja y bancos, transferencias, libro mayor y dashboard operativo",
                        "movimiento bancario",
                        "el detalle del movimiento",
                        "consulta de caja y bancos",
                        "saldo bancario",
                        "cuenta corriente y operacion",
                        "el periodo abierto")),
                    CrearModulo("TRANSFERENCIACUENTA", "Transferencias", "bi-arrow-left-right", "Traslado entre cuentas corrientes con doble seccion emisor y receptor.", CrearFaqRegistro(
                        "Transferencias",
                        "registrar el egreso de una cuenta y el ingreso correlativo en otra cuenta corriente",
                        "dos cuentas corrientes activas, periodo abierto, fecha y tipo de cambio consistente entre ambas secciones",
                        "cuenta emisora, cuenta receptora, operacion, fecha, monto, glosa y observacion",
                        "la transferencia mueve saldos entre cuentas sin alterar el total de tesoreria de la empresa",
                        "usar la misma cuenta en ambos lados, no cuadrar conversion o dejar una seccion incompleta",
                        "caja y bancos, saldos por cuenta, libro mayor bancario y conciliacion operativa",
                        "consultar movimientos de ambas cuentas y comprobar que se registraron egreso e ingreso enlazados",
                        "tesoreria o responsable de bancos",
                        "caja y bancos, libro mayor, dashboard y conciliacion",
                        "transferencia entre cuentas",
                        "las dos secciones emisor y receptor",
                        "consulta de caja y bancos",
                        "saldo por cuenta",
                        "cuenta emisora y receptora",
                        "el periodo abierto")),
                    CrearModulo("APLICACION", "Aplicaciones", "bi-link-45deg", "Compensacion de comprobantes con notas de credito.", CrearFaqRegistro(
                        "Aplicaciones",
                        "compensar comprobantes pendientes con notas de credito del mismo cliente o proveedor",
                        "periodo abierto, persona seleccionada, documento pendiente y nota de credito disponible compatible",
                        "tipo de persona, fecha, moneda, importe aplicado, glosa, comprobante pendiente y nota de credito",
                        "la aplicacion reduce saldos abiertos y genera trazabilidad de compensacion documental",
                        "aplicar una NC a una persona distinta, mezclar monedas sin TC o exceder el saldo disponible",
                        "saldos de compras o ventas, cuenta corriente del tercero y control del asiento APNC",
                        "revisar el saldo del comprobante y de la NC despues de guardar la aplicacion",
                        "cuentas por cobrar o pagar con soporte contable",
                        "compras, ventas, analisis por auxiliar, libro mayor y dashboard",
                        "aplicacion de nota de credito",
                        "la relacion comprobante y NC",
                        "consulta de aplicaciones",
                        "saldo compensado",
                        "persona, comprobante y NC",
                        "el periodo abierto"))
                ]),
            CrearCategoria(
                "PROCESO",
                "Proceso",
                "bi-gear-wide-connected",
                "Generaciones mensuales o anuales que recalculan, ajustan o cierran la contabilidad.",
                [
                    CrearModulo("DIFERENCIACAMBIO", "Diferencia en Cambio", "bi-currency-exchange", "Proceso mensual para cuentas monetarias en moneda extranjera.", CrearFaqProceso(
                        "Diferencia en Cambio",
                        "generar el ajuste por variacion de tipo de cambio sobre saldos monetarios",
                        "periodo abierto, tipos de cambio del periodo y reglas de cuentas destino o contrapartida",
                        "cuentas monetarias, saldos en moneda extranjera, TC de cierre y cuentas de ganancia o perdida",
                        "cuando ya registraste operaciones del mes y necesitas reexpresar saldos al cierre mensual",
                        "correrlo sin tipos de cambio completos o sin validar cuentas destino de ganancia y perdida",
                        "asientos de ajuste del periodo, balance, libro mayor y cierre mensual",
                        "revisar detalle generado, asientos emitidos y efecto en saldos antes de cerrar periodo",
                        "contabilidad financiera al cierre mensual",
                        "libro mayor, balance, cierre y estados financieros")),
                    CrearModulo("AJUSTECUENTA", "Ajuste de Cuentas", "bi-sliders", "Proceso por cuenta analitica con configuracion de cuenta destino.", CrearFaqProceso(
                        "Ajuste de Cuentas",
                        "generar asientos AJU por cuenta analitica segun la configuracion funcional de destino",
                        "periodo abierto, configuracion de cuentas destino y cuentas analiticas con saldos a procesar",
                        "cuenta analitica, cuenta destino, signo, importe y origen de ajuste",
                        "cuando debes reclasificar resultados o distribuir saldos a cuentas finales del periodo",
                        "procesar sin revisar configuracion de contrapartidas o ejecutar sobre cuentas sin analisis correcto",
                        "asientos de ajuste, libro mayor, balance y consistencia del cierre mensual",
                        "consultar el detalle previo, generar y comparar antes y despues del saldo de cada cuenta",
                        "contabilidad funcional o responsable del cierre",
                        "balance, libro mayor, cierre y reportes de resultado")),
                    CrearModulo("CERRARPERIODO", "Cerrar Periodo", "bi-lock", "Bloqueo operativo de un periodo para evitar nuevas modificaciones.", CrearFaqProceso(
                        "Cerrar Periodo",
                        "bloquear un mes para impedir nuevos registros una vez validada la contabilidad",
                        "haber ejecutado conciliaciones, ajustes, validaciones y revisiones de saldos del mes",
                        "periodo, estado abierto o cerrado, observaciones de cierre y control de autorizacion",
                        "cuando compras, ventas, bancos y asientos ya fueron revisados y no deben alterarse",
                        "cerrar un periodo con pendientes de validacion, saldos abiertos criticos o procesos sin correr",
                        "todos los registros del mes, reportes oficiales y la disciplina de cierres cronologicos",
                        "verificar el estado del periodo y probar que los formularios dejan de permitir grabacion",
                        "contabilidad con autorizacion del responsable del cierre",
                        "todos los registros del periodo, reportes y aperturas o cierres siguientes")),
                    CrearModulo("ASIENTOCIERRE", "Asiento de cierre", "bi-journal-x", "Generacion anual de un unico asiento compuesto en el periodo 14 que invierte los saldos de las cuentas configuradas como Inventario.", CrearFaqProceso(
                        "Asiento de cierre",
                        "emitir un unico asiento que invierte los saldos acumulados y deja las cuentas listas para el siguiente ejercicio",
                        "el origen CIE configurado, tipo de cambio USD del 31/12, periodo de corte entre 00 y 13, ajustes ejecutados y saldos validados en mayor y balance",
                        "las cuentas con saldo y ColBalance I - Inventario, sus importes en soles y dolares, el sentido Debe/Haber y el origen de cierre",
                        "cuando el ejercicio esta completo y se aprobo la informacion de cierre anual",
                        "cerrar con diferencias pendientes o sin revisar procesos previos como ajuste y diferencia en cambio",
                        "libro diario, libro mayor, balance anual y asiento de apertura siguiente",
                        "revisar cada linea invertida, comparar los totales Debe y Haber aunque sean diferentes y confirmar el asiento generado antes de la apertura",
                        "contabilidad de cierre anual",
                        "balance anual, libro mayor, asiento de apertura y libros electronicos",
                        ("En que periodo se genera el asiento de cierre", "Se genera obligatoriamente en el periodo 14 - Cierre de Inventario. El usuario solo elige hasta que periodo acumular saldos, desde 00 hasta 13."),
                        ("Que cuentas incluye el asiento de cierre", "Incluye unicamente cuentas del plan empresarial con ColBalance I - Inventario cuyo saldo acumulado absoluto sea al menos 0.01 entre el periodo 00 y el corte seleccionado."),
                        ("Como invierte los saldos el asiento de cierre", "Un saldo deudor se registra en el Haber y un saldo acreedor en el Debe. Los importes reales en soles y dolares se conservan por linea."),
                        ("Por que el asiento de cierre puede quedar descuadrado", "El proceso no agrega una cuenta de contrapartida ni una linea artificial de cuadre. Por diseño, el total Debe puede ser diferente del total Haber."),
                        ("Que tipo de cambio utiliza el cierre", "Requiere un tipo de cambio USD activo al 31/12. Usa Compra y Venta regulares o CompraSBS y VentaSBS segun el parametro empresarial TIPO_CAMBIO_SBS_CIERRE, y permite revisar los valores antes de generar."),
                        ("Que sucede al regenerar el asiento de cierre", "El sistema elimina la generacion CIE anterior del mismo ejercicio, recalcula los correlativos afectados y crea nuevamente el asiento compuesto con el corte y tipos de cambio actuales."),
                        ("Que elimina el boton Eliminar asiento de cierre", "Elimina exclusivamente el proceso CIE del ejercicio, su detalle y los asientos vinculados; despues recompone o elimina los correlativos que correspondan."),
                        ("Que hago si aparece una generacion anterior con varios asientos", "Usa Regenerar asiento de cierre. El proceso reconoce el modelo anterior, elimina sus asientos vinculados y los reemplaza por el nuevo asiento compuesto."),
                        ("Como reviso el asiento generado desde la pantalla de cierre", "Consulta el ejercicio y abre el numero de asiento mostrado. El detalle presenta cada cuenta, moneda, sentido Debe/Haber, tipo de cambio e importes en soles y dolares."),
                        ("Que ocurre si no existen cuentas de Inventario con saldo", "El proceso no genera un asiento y muestra que no existen cuentas configuradas como Inventario con saldo pendiente para cerrar."))),
                    CrearModulo("ASIENTOAPERTURA", "Asiento de apertura", "bi-journal-plus", "Generacion anual del asiento inicial del ejercicio.", CrearFaqProceso(
                        "Asiento de apertura",
                        "crear el asiento inicial del nuevo ejercicio a partir de saldos finales del periodo anterior",
                        "cierre del ejercicio previo, origen configurado y cuentas listas para apertura anual",
                        "periodo destino, cuenta, saldo inicial, cuenta de resultado y origen de apertura",
                        "al iniciar un nuevo ejercicio contable luego del cierre anual validado",
                        "ejecutarlo sin cierre previo correcto o duplicar una apertura existente del mismo ejercicio",
                        "saldos iniciales del libro mayor, balance y reportes del nuevo periodo",
                        "revisar el detalle de apertura, el asiento generado y los saldos iniciales del libro mayor",
                        "contabilidad general en el inicio del ejercicio",
                        "libro mayor, balance, cierre anual y reportes del nuevo anio"))
                ]),
            CrearCategoria(
                "REPORTES",
                "Reportes",
                "bi-bar-chart",
                "Consultas de control, emision A4 y libros oficiales para revision y cierre contable.",
                [
                    CrearModulo("ANALISISCUENTAS", "Analisis de cuentas", "bi-bar-chart", "Seguimiento por cuenta, persona y documento.", CrearFaqReporte(
                        "Analisis de cuentas",
                        "seguir movimientos por cuenta, auxiliar, documento y vista detallada o resumida",
                        "periodo, estado, documento, vista y rango de cuentas seleccionado desde la ayuda emergente",
                        "rastrear composicion de saldos y movimientos por tercero o documento",
                        "cuando necesitas explicar un saldo o revisar una cuenta analitica",
                        "validar auxiliares, referencias y glosas frente al libro mayor",
                        "contabilidad, auditoria interna y revision operativa",
                        "detalle por cuenta, persona, documento y glosa",
                        "comprobacion de saldos, soportes y trazabilidad documental")),
                    CrearModulo("LIBRODIARIO", "Libro Diario", "bi-journal-richtext", "Consulta del diario auxiliar, por cuenta y por origen.", CrearFaqReporte(
                        "Libro Diario",
                        "consultar asientos del periodo en orden cronologico, por cuenta u origen",
                        "anio, periodo, vista y rango opcional de cuentas; la ayuda permite elegir niveles 1 al 5",
                        "revisar el libro diario del mes y validar correlativos u origenes",
                        "cuando auditas asientos manuales y automaticos del periodo",
                        "comparar contra asientos, origenes y glosas del movimiento real",
                        "contabilidad, auditoria y cierre mensual",
                        "asientos del periodo con debe, haber y referencia del origen",
                        "soporte cronologico del libro diario")),
                    CrearModulo("LIBROMAYOR", "Libro Mayor", "bi-journals", "Mayor por cuenta con saldo inicial, movimientos y cierre.", CrearFaqReporte(
                        "Libro Mayor",
                        "revisar el movimiento y saldo final de cada cuenta contable",
                        "anio, mes, rango de cuentas mediante la ayuda emergente y documento si se requiere detalle",
                        "explicar saldos por cuenta y analizar acumulados del periodo",
                        "cuando necesitas validar el comportamiento de una cuenta especifica",
                        "comparar saldo inicial, movimientos, subtotales y saldo final con balance",
                        "contabilidad financiera y control interno",
                        "saldo anterior, movimientos, subtotal y saldo final por cuenta",
                        "validacion de saldos del periodo")),
                    CrearModulo("REGISTROVENTAS", "Registro de ventas", "bi-receipt-cutoff", "Formato mensual de ventas para control y emision.", CrearFaqReporte(
                        "Registro de ventas",
                        "emitir el reporte mensual de ventas en formato de control contable",
                        "anio, periodo, DNI/RUC del cliente y numero del comprobante; la ayuda de persona completa el DNI/RUC",
                        "presentar y revisar los comprobantes de ventas del mes",
                        "cuando debes validar la provision de ventas o preparar un soporte externo",
                        "comparar totales, impuestos y series contra las ventas registradas",
                        "contabilidad, impuestos y cierre mensual",
                        "comprobantes, clientes, bases imponibles, impuestos y totales",
                        "control mensual de ventas registradas")),
                    CrearModulo("REGISTROCOMPRAS", "Registro de compras", "bi-cart3", "Formato mensual de compras, IGV y saldos de provision.", CrearFaqReporte(
                        "Registro de compras",
                        "emitir el reporte mensual de compras segun las provisiones registradas",
                        "anio, periodo, DNI/RUC del proveedor y numero del comprobante; la ayuda de persona completa el DNI/RUC",
                        "controlar comprobantes de compras, impuestos y saldos del mes",
                        "cuando revisas IGV, proveedores o consistencia del registro tributario",
                        "comparar documento, proveedor, base imponible, IGV y total con la provision",
                        "contabilidad, impuestos y cuentas por pagar",
                        "comprobantes, proveedores, bases, IGV y totales del periodo",
                        "control mensual de compras provisionadas")),
                    CrearModulo("BALANCECOMPROBACION", "Balance de comprobacion", "bi-table", "Saldos y movimientos por grado del plan contable.", CrearFaqReporte(
                        "Balance de comprobacion",
                        "mostrar saldos y movimientos por cuentas y grados del plan contable",
                        "anio, periodos, moneda, grado y rango de cuentas de cualquier nivel; Todas las cuentas desactiva el rango",
                        "validar el cuadre general del periodo y navegar por niveles del plan",
                        "cuando necesitas una vista resumida o detallada de saldos antes del cierre",
                        "comparar debe, haber, saldos, inventario y resultados por naturaleza o funcion",
                        "contabilidad financiera y revision de cierre",
                        "cuentas por grado, movimientos, saldos y agrupaciones del balance",
                        "cuadre general del periodo y revision de estructura")),
                    CrearModulo("LIBROELECTRONICO", "Libros Electronicos", "bi-filetype-txt", "Consulta, validacion, generacion, descarga y control de presentacion de archivos TXT para el PLE.", CrearFaqLibrosElectronicos())
                ])
        ];
    }

    private static AyudaCategoriaViewModel CrearCategoria(
        string clave,
        string titulo,
        string icono,
        string descripcion,
        IReadOnlyCollection<AyudaModuloViewModel> modulos)
    {
        return new AyudaCategoriaViewModel
        {
            Clave = clave,
            Titulo = titulo,
            Icono = icono,
            Descripcion = descripcion,
            Modulos = modulos
        };
    }

    private static AyudaModuloViewModel CrearModulo(
        string clave,
        string titulo,
        string icono,
        string resumen,
        IReadOnlyCollection<AyudaPreguntaViewModel> preguntas)
    {
        return new AyudaModuloViewModel
        {
            Clave = clave,
            Titulo = titulo,
            Icono = icono,
            Resumen = resumen,
            Preguntas = preguntas
        };
    }

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqDashboard()
        => CrearPreguntas("dashboard",
        [
            ("Para que sirve el dashboard", "Resume el periodo activo con KPIs de compras, ventas, asientos, cuentas corrientes e indicadores operativos para detectar rapidamente pendientes y volumen de trabajo."),
            ("Que informacion debo revisar primero", "Empieza por el estado del panel y del periodo, luego valida compras, ventas y asientos del mes, y finalmente los indicadores de pendientes."),
            ("Que muestran los KPIs principales", "Muestran cantidad e importe de compras, ventas y asientos del periodo, ademas de las cuentas corrientes activas disponibles para operacion."),
            ("Que significa el estado del panel", "Es una lectura de la vigencia de la cuenta administradora y de la suscripcion asociada a la empresa activa."),
            ("Que significa periodo abierto o cerrado", "Periodo abierto permite registrar y modificar operaciones. Periodo cerrado bloquea grabaciones nuevas y deja el mes solo para consulta."),
            ("Como interpreto los indicadores de control", "Te senalan compras o ventas con saldo, compras sin asiento y validaciones CPE ya realizadas para priorizar seguimiento."),
            ("Para que sirven los graficos de torres", "Comparan importes PEN, USD y cantidad de movimientos por periodo para compras, ventas y movimientos bancarios, lo que permite ver tendencias y picos mensuales."),
            ("Para que sirve la torta de distribucion del periodo", "Mide la participacion de compras, ventas, asientos, movimientos bancarios y aplicaciones dentro del volumen operativo del mes."),
            ("Que hago si el dashboard no refleja una operacion reciente", "Verifica que estes en la empresa y periodo correctos, que la operacion se haya guardado y que no exista un filtro operativo pendiente en su modulo de origen."),
            ("Cada cuanto conviene revisar el dashboard", "Al inicio del dia, antes del cierre de jornada y previo a ejecutar procesos mensuales o bloquear un periodo.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqEmpresas()
        => CrearPreguntas("empresas",
        [
            ("Para que sirve la opcion Empresas", "Permite elegir la empresa activa de trabajo o registrar una nueva empresa dentro de la cuenta administradora."),
            ("Que pasa si no selecciono una empresa activa", "El panel administrativo redirige a la seleccion de empresa porque todos los mantenimientos y registros se ejecutan sobre una empresa concreta."),
            ("Quien debe crear una nueva empresa", "El usuario con autorizacion administrativa de la cuenta, ya que la empresa queda vinculada a la suscripcion y al contexto contable."),
            ("Que datos debo validar al registrar una empresa", "Razon social, nombre comercial, RUC, correo principal, vigencia y datos de la cuenta administradora a la que quedara asociada."),
            ("Puedo corregir los datos de una empresa registrada", "Si. Desde Seleccion de empresa usa Editar para corregir la razon social, el nombre comercial o el RUC. Al guardar, el codigo interno se sincroniza con el nuevo RUC y permanece oculto en el selector."),
            ("Cambiar de empresa altera mis datos actuales", "No. Solo cambia el contexto de consulta y registro; cada empresa conserva sus propios catalogos y movimientos."),
            ("Como saber en que empresa estoy trabajando", "El nombre aparece en la tarjeta Empresa activa del menu lateral y tambien se refleja en los encabezados del panel."),
            ("Puedo usar el mismo usuario en varias empresas", "Si, siempre que el usuario este vinculado a la cuenta administradora y tenga acceso a esas empresas."),
            ("Que impacto tiene una empresa nueva en la suscripcion", "Consume capacidad del limite de empresas permitidas definido para la cuenta administradora cuando aplica un tope."),
            ("Una empresa nueva recibe automaticamente el plan y la configuracion contable", "No. El registro de la empresa no dispara cargas maestras. La configuracion inicial debe ejecutarse expresamente desde los mantenimientos autorizados."),
            ("Como inicializo contablemente una empresa nueva", "Primero usa Cargar configuracion contable por defecto desde Plan de cuentas para crear plan, parametros, cuentas destino, impuestos y documentos. Luego usa Cargar origenes por defecto para crear los subdiarios y su configuracion contable inicial."),
            ("Que hago si una empresa no aparece en el selector", "Verifica que el usuario tenga acceso, que la empresa este activa y que la vinculacion con la cuenta administradora exista."),
            ("Cuanto conviene revisar esta pantalla", "Cada vez que cambies de contexto de trabajo, al dar de alta una empresa o cuando necesites confirmar sobre que empresa operaras el periodo.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqUsuarios()
        => CrearPreguntas("usuarios",
        [
            ("Para que sirve la opcion Usuarios", "Permite crear o vincular usuarios a la cuenta administradora, asignarles un rol y definir en que empresas pueden trabajar."),
            ("Que datos necesito para registrar un usuario", "Debes indicar nombre completo, correo, telefono, rol de cuenta y seleccionar al menos una empresa disponible."),
            ("Cuando debo ingresar una contrasena temporal", "La contrasena temporal es obligatoria cuando el correo aun no pertenece a un usuario del sistema. Al iniciar sesion, el nuevo usuario debera cambiarla."),
            ("Que ocurre si el correo ya pertenece a un usuario", "El sistema reutiliza la cuenta existente y la vincula a la cuenta administradora, evitando crear un acceso duplicado para el mismo correo."),
            ("Por que no puedo agregar otro usuario", "La cuenta puede haber alcanzado el limite de usuarios permitido por la suscripcion. Revisa Mi suscripcion antes de intentar una nueva alta."),
            ("Para que sirve el rol de cuenta", "El rol establece el nivel base de acceso del usuario dentro de la cuenta administradora y sirve como referencia para calcular sus permisos efectivos."),
            ("Como asigno o retiro empresas a un usuario", "Desde la opcion Permisos puedes marcar las empresas donde trabajara el usuario. Debe conservar al menos una empresa asignada."),
            ("Que diferencia hay entre permisos de cuenta y permisos por empresa", "Los permisos de cuenta aplican como excepciones generales; los permisos por empresa permiten ajustar el acceso operativo solo dentro de una empresa seleccionada."),
            ("Que significan Ver, Crear, Editar y Eliminar en Permisos", "Son las acciones controladas por modulo. La opcion Rol hereda el valor base y Si o No crea una excepcion explicita para esa accion."),
            ("Que pasa cuando desactivo el acceso de un usuario", "Se desactiva su vinculacion con la cuenta administradora sin eliminar su identidad ni el historial de operaciones que ya registro.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqConfiguracion()
        => CrearPreguntas("configuracion",
        [
            ("Para que sirve la opcion Configuracion", "Permite guardar los datos de facturacion de la cuenta administradora que se reutilizaran al emitir sus comprobantes comerciales."),
            ("Que comprobante puedo elegir como preferido", "Puedes seleccionar Boleta o Factura. Esta preferencia determina que datos fiscales deben quedar completos antes de guardar."),
            ("Que necesito para configurar una factura", "Debes seleccionar RUC como tipo de documento, ingresar un RUC valido de 11 digitos y completar la razon social de facturacion."),
            ("Que necesito para configurar una boleta", "Debes completar el nombre de facturacion y registrar el tipo y numero de documento que correspondan al titular."),
            ("Para que sirve el boton Consultar", "Consulta el padron externo usando el documento ingresado y carga los datos encontrados para reducir la digitacion manual."),
            ("Con que documentos funciona la consulta automatica", "La consulta automatica esta disponible para DNI y RUC. Los demas tipos de documento deben completarse manualmente."),
            ("Cuantos digitos deben tener el DNI y el RUC", "El DNI debe contener 8 digitos y el RUC 11 digitos. El formulario impide guardar cuando estas longitudes no son validas."),
            ("Que datos puede completar la consulta al padron", "Puede completar nombre o razon social, direccion fiscal, ubigeo, distrito, provincia y departamento cuando el servicio devuelve esa informacion."),
            ("Puedo corregir los datos obtenidos del padron", "Si. Revisa y corrige los campos antes de guardar porque la informacion almacenada sera la referencia de facturacion de la cuenta."),
            ("Que debo revisar antes de guardar la configuracion", "Confirma el tipo de comprobante, documento, nombre o razon social, direccion fiscal y ubicacion. Atiende cualquier mensaje de validacion antes de continuar.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqMiSuscripcion()
        => CrearPreguntas("suscripcion",
        [
            ("Que informacion muestra Mi suscripcion", "Estado del plan, tipo de plan, fechas de vigencia, limites operativos y datos de la cuenta administradora que controla la empresa activa."),
            ("De donde sale esta informacion", "Del contexto de suscripcion asociado a la empresa activa y a la cuenta administradora vinculada en seguridad."),
            ("Que significa estado TRIAL", "Que la empresa o cuenta administradora se encuentra operando en una etapa de prueba sujeta a fechas de inicio y fin configuradas."),
            ("Que significa estado ACTIVO", "Que la suscripcion se encuentra vigente y con operacion habilitada conforme a su plan contratado."),
            ("Que significa estado SUSPENDIDO o BAJA", "Que existe una restriccion o termino de la suscripcion y deben revisarse vigencia, pagos o autorizaciones administrativas."),
            ("Como se calcula la fecha de vencimiento visible", "Si la cuenta esta en prueba se toma FechaFinPrueba; en caso contrario se usa FechaFinPlan."),
            ("Para que sirven los limites de empresas y usuarios", "Indican la capacidad contratada para crecer en numero de empresas asociadas y usuarios activos dentro de la cuenta administradora."),
            ("Que hago si la suscripcion esta por vencer", "Coordina renovacion antes del vencimiento y evita ejecutar cierres o aperturas sensibles sin confirmar continuidad del servicio."),
            ("Que hago si los datos de contacto no son correctos", "Escala la observacion al administrador de plataforma o al responsable interno que gestione la cuenta administradora."),
            ("Cuando debo revisar esta pantalla", "Siempre que exista una alerta de acceso, antes de alta de nuevas empresas o usuarios y como control preventivo del estado contractual.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqAyuda()
        => CrearPreguntas("ayuda",
        [
            ("Para que sirve la opcion Ayuda", "Centraliza preguntas y respuestas operativas de cada modulo del sistema para estandarizar el uso diario."),
            ("Que es la ayuda contextual", "Es el acceso directo Ayuda de este modulo, que abre esta pantalla enfocada en la opcion desde la que estabas navegando."),
            ("Como encuentro rapido una respuesta", "Usa el buscador de preguntas dentro de la categoria o del modulo visible para filtrar por terminos como saldo, asiento, cierre o proveedor."),
            ("La ayuda cambia segun el modulo", "Si. Cada controlador principal y cada opcion de reportes o procesos tiene su propio bloque de preguntas."),
            ("Puedo entrar a Ayuda desde el menu general", "Si. Tambien puedes abrirla desde el acceso contextual que aparece dentro del panel administrativo."),
            ("La ayuda explica solo la pantalla o tambien el criterio contable", "Incluye ambos enfoques: uso de la interfaz y criterio funcional asociado a saldos, documentos, procesos y reportes."),
            ("Que hago si no encuentro una pregunta exacta", "Revisa la categoria del modulo relacionado y usa palabras clave cercanas al proceso real que estas ejecutando."),
            ("La ayuda reemplaza la documentacion funcional", "No. La complementa con respuestas operativas rapidas basadas en la estructura y flujos del sistema."),
            ("Puedo usar Ayuda durante el cierre mensual", "Si, y es recomendable para recordar prerrequisitos de cada proceso y de los reportes de validacion."),
            ("Cada cuanto conviene revisar o ampliar la ayuda", "Cuando se agregan modulos, cambian reglas de negocio o el equipo detecta dudas repetidas en operacion.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqMantenimiento(
        string modulo,
        string proposito,
        string prerrequisitos,
        string camposClave,
        string cuandoCrear,
        string errorComun,
        string impacto,
        string validacion,
        string responsable,
        string dependencias,
        params (string pregunta, string respuesta)[] preguntasAdicionales)
    {
        var preguntas = new List<(string pregunta, string respuesta)>
        {
            ($"Para que sirve {modulo}", $"Sirve para {proposito}."),
            ($"Que debo tener listo antes de usar {modulo}", $"Antes de usar {modulo} conviene validar {prerrequisitos}."),
            ($"Que campos son los mas importantes en {modulo}", $"Los campos clave de {modulo} son {camposClave}."),
            ($"Cuando corresponde crear un nuevo registro en {modulo}", $"Debes crear uno nuevo {cuandoCrear}."),
            ($"Cuando es mejor desactivar y no eliminar en {modulo}", "Cuando el registro ya tiene historial operativo. Desactivar evita romper trazabilidad y mantiene limpio el uso futuro."),
            ($"Que error comun debo evitar en {modulo}", $"El error mas comun es {errorComun}."),
            ($"En que impacta {modulo} dentro del sistema", $"El impacto principal de {modulo} esta en {impacto}."),
            ($"Como valido que el registro de {modulo} quedo correcto", $"La validacion recomendada es {validacion}."),
            ($"Quien deberia mantener {modulo}", $"{modulo} deberia ser administrado por {responsable}."),
            ($"Que otros modulos dependen de {modulo}", $"Las dependencias principales de {modulo} son {dependencias}.")
        };

        preguntas.AddRange(preguntasAdicionales);
        return CrearPreguntas(NormalizarClave(modulo).ToLowerInvariant(), preguntas);
    }

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqRegistro(
        string modulo,
        string proposito,
        string prerrequisitos,
        string camposClave,
        string impactoSaldo,
        string errorComun,
        string consulta,
        string validacion,
        string responsable,
        string dependencias,
        string registroNombre,
        string detalleNombre,
        string reportePrincipal,
        string conceptoSaldo,
        string actorPrincipal,
        string requisitoPeriodo)
        => CrearPreguntas(NormalizarClave(modulo).ToLowerInvariant(),
        [
            ($"Para que sirve {modulo}", $"Sirve para {proposito}."),
            ($"Que debo revisar antes de registrar en {modulo}", $"Antes de registrar en {modulo} revisa {prerrequisitos}."),
            ($"Que datos son obligatorios o criticos en {modulo}", $"En {modulo} es clave controlar {camposClave}."),
            ($"Como impacta {modulo} en saldos y contabilidad", $"El impacto principal es que {impactoSaldo}."),
            ($"Cual es el error mas comun en {modulo}", $"Un error frecuente es {errorComun}."),
            ($"Que pantalla de consulta debo revisar despues de guardar en {modulo}", $"Despues de guardar conviene revisar {consulta}."),
            ($"Como valido que el registro se guardo bien en {modulo}", $"La validacion recomendada es {validacion}."),
            ($"Quien deberia operar {modulo}", $"{modulo} deberia ser operado por {responsable}."),
            ($"Que otros modulos se cruzan con {modulo}", $"Las dependencias principales de {modulo} son {dependencias}."),
            ($"Que representa exactamente un {registroNombre}", $"Representa la unidad operativa principal que se registra desde el modulo {modulo}."),
            ($"Para que sirve {detalleNombre} en {modulo}", $"{detalleNombre} permite desagregar cuentas, documentos, impuestos o conceptos segun el circuito del modulo."),
            ($"Que reporte refleja normalmente la informacion de {modulo}", $"El reporte mas relacionado con {modulo} es {reportePrincipal}."),
            ($"Que significa controlar el {conceptoSaldo} en {modulo}", "Significa confirmar que el registro deja correctamente valorizada la obligacion o derecho generado por la operacion."),
            ($"Por que es importante validar {actorPrincipal} en {modulo}", $"Porque el registro queda trazado contra {actorPrincipal} y eso afecta consultas por auxiliar, documento o saldos abiertos."),
            ($"Que pasa si el periodo no esta disponible al registrar", $"No deberias registrar. {modulo} exige {requisitoPeriodo} para mantener la secuencia cronologica del sistema."),
            ($"Que debo revisar si el boton Guardar no deberia usarse aun", "Revisa campos obligatorios, cuadre de montos, periodo vigente y que los selectores principales tengan valores validos."),
            ($"Cuando conviene editar un registro existente de {modulo}", "Solo cuando el periodo siga abierto y la correccion no rompa documentos, saldos o procesos posteriores ya ejecutados."),
            ($"Que hago si el registro necesita anulacion o reemplazo", "Evalua si corresponde editar, revertir mediante otro movimiento o dejar trazabilidad con un nuevo registro compensatorio segun el circuito."),
            ($"Como detecto duplicados en {modulo}", $"Busca por periodo, {actorPrincipal} y referencias como documento, serie, numero o glosa antes de crear un nuevo registro."),
            ($"Que revision final conviene hacer al terminar de trabajar en {modulo}", "Confirma el resultado en su consulta principal, revisa el saldo o total generado y valida el impacto en reportes o asientos relacionados.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqProceso(
        string modulo,
        string proposito,
        string prerrequisitos,
        string datosClave,
        string cuandoEjecutar,
        string errorComun,
        string impacto,
        string validacion,
        string responsable,
        string dependencias,
        params (string pregunta, string respuesta)[] preguntasAdicionales)
    {
        var preguntas = new List<(string pregunta, string respuesta)>
        {
            ($"Para que sirve {modulo}", $"Sirve para {proposito}."),
            ($"Que debo verificar antes de ejecutar {modulo}", $"Antes de ejecutar {modulo} valida {prerrequisitos}."),
            ($"Que datos son determinantes en {modulo}", $"Los datos o parametros mas sensibles en {modulo} son {datosClave}."),
            ($"Cuando corresponde correr {modulo}", $"{modulo} debe ejecutarse {cuandoEjecutar}."),
            ($"Que error comun debo evitar en {modulo}", $"El error mas comun es {errorComun}."),
            ($"En que impacta {modulo} dentro del cierre", $"El impacto principal de {modulo} esta en {impacto}."),
            ($"Que modulos dependen de que {modulo} salga bien", $"Las dependencias principales son {dependencias}."),
            ($"Como valido el resultado de {modulo}", $"La validacion recomendada es {validacion}."),
            ($"Quien deberia ejecutar {modulo}", $"{modulo} deberia ser ejecutado por {responsable}."),
            ($"Puedo volver a ejecutar {modulo} si cambio informacion previa", "Solo si el proceso y su pantalla lo permiten. Antes revisa si elimina la generacion previa o si necesita reversa manual para no duplicar efectos."),
            ($"Que hago si {modulo} genera un resultado inesperado", "Deten el cierre, revisa configuraciones, saldos fuente y dependencias del periodo antes de volver a correr el proceso."),
            ($"En que orden revisar los resultados de {modulo}", "Primero el detalle generado, luego los asientos asociados y finalmente el impacto en reportes como libro mayor y balance."),
            ($"Que revision final conviene hacer despues de {modulo}", "Comparar antes y despues del saldo afectado y dejar el periodo listo para el siguiente proceso del cierre.")
        };

        preguntas.AddRange(preguntasAdicionales);
        return CrearPreguntas(NormalizarClave(modulo).ToLowerInvariant(), preguntas);
    }

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqReporte(
        string modulo,
        string proposito,
        string filtros,
        string utilidad,
        string cuandoUsar,
        string contraste,
        string responsable,
        string salida,
        string control)
        => CrearPreguntas(NormalizarClave(modulo).ToLowerInvariant(),
        [
            ($"Para que sirve {modulo}", $"Sirve para {proposito}."),
            ($"Que filtros debo revisar en {modulo} antes de consultar", $"Conviene revisar {filtros}."),
            ($"Que utilidad principal tiene {modulo}", $"La utilidad principal de {modulo} es {utilidad}."),
            ($"Cuando conviene usar {modulo}", $"Conviene usar {modulo} {cuandoUsar}."),
            ($"Contra que deberia contrastar la informacion de {modulo}", $"Lo recomendable es contrastar {modulo} con {contraste}."),
            ($"Quien usa normalmente {modulo}", $"{modulo} es usado normalmente por {responsable}."),
            ($"Que salida muestra {modulo}", $"{modulo} muestra {salida}."),
            ($"Que control me ayuda a hacer {modulo}", $"Te ayuda a realizar {control}."),
            ($"Que hago si el reporte sale vacio", "Revisa periodo, empresa activa, filtros de cuenta o documento y confirma que existan registros en el modulo fuente."),
            ($"Que hago si los totales no cuadran como esperaba", "Compara el reporte con los movimientos base, verifica moneda, rango de cuentas y si el periodo ya incluye ajustes o cierres."),
            ($"Puedo usar {modulo} antes de cerrar el periodo", $"Si, y de hecho {modulo} es util como control previo para detectar diferencias antes del cierre definitivo."),
            ($"Que revision final conviene hacer al terminar con {modulo}", "Guardar o imprimir la consulta necesaria, dejar evidencia del corte revisado y pasar al siguiente control del cierre o de la operacion.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearFaqLibrosElectronicos()
        => CrearPreguntas("libroselectronicos",
        [
            ("Para que sirve Libros Electronicos", "Permite consultar, validar, generar y descargar los archivos TXT requeridos por el PLE para los formatos 5.1 Libro Diario, 5.2 Libro Diario Simplificado y 6.1 Libro Mayor."),
            ("Que debo seleccionar antes de consultar", "Verifica la empresa activa, el anio, el mes y el libro electronico. La exportacion se genera en soles para el periodo seleccionado."),
            ("Que hace ahora el boton Consultar", "Consultar carga el resumen y el detalle exportable, ejecuta automaticamente las validaciones y deja disponibles las Observaciones del PLE en la misma pantalla."),
            ("Que revisa la validacion automatica", "Revisa RUC, periodo, cuadre de asientos y totales, CUO y correlativos, cuentas, monedas, documentos, comprobantes, importes, glosas y estados admitidos por el formato PLE."),
            ("Como interpreto las Observaciones del PLE", "Los errores bloquean la generacion. Las advertencias requieren revision y los mensajes informativos explican condiciones validas, como un periodo cerrado o un periodo sin movimientos."),
            ("Que archivos genera el boton Generar TXT", "Para el Libro Diario 5.1 genera tambien el plan contable 5.3; para el Diario Simplificado 5.2 genera tambien el plan 5.4. Para el Libro Mayor 6.1 genera solamente su archivo principal."),
            ("Por que debo descargar dos TXT para los diarios", "Los formatos 5.1 y 5.2 son libros compuestos. Debes descargar el archivo principal y su plan contable 5.3 o 5.4, y seleccionarlos simultaneamente en el PLE."),
            ("Cuando el plan 5.3 o 5.4 se genera completo", "Se genera completo en la primera presentacion del ejercicio o cuando no existe un snapshot anterior presentado para la misma empresa y el mismo libro."),
            ("Cuando el plan 5.3 o 5.4 contiene solo cambios", "Despues de una presentacion confirmada, el sistema compara el plan actual con el ultimo snapshot presentado y exporta solo cuentas nuevas o aquellas cuyo codigo o nombre fue modificado."),
            ("Que ocurre si el plan contable no tuvo cambios", "El archivo 5.3 o 5.4 se genera vacio y su nombre usa el indicador de contenido 0. Aunque este vacio, debe cargarse junto con el archivo principal en el PLE."),
            ("Puedo generar un libro de un periodo sin movimientos", "Si. La ausencia de asientos y lineas se informa como Periodo sin movimientos y se genera el TXT principal vacio con indicador de contenido 0, sin tratarlo como error."),
            ("Para que sirve Marcar como presentado", "Debes activarlo solo despues de obtener la constancia de recepcion del PLE. La marca confirma la presentacion y convierte el snapshot generado en la referencia para comparar el plan de los meses siguientes."),
            ("Por que puede bloquearse la generacion de un mes", "Si el mes anterior tiene movimientos y todavia no fue marcado como presentado, el sistema bloquea el siguiente mes para conservar la secuencia de presentacion."),
            ("Puedo desmarcar o volver a generar un periodo", "Puedes desmarcar y volver a generar mientras no exista un periodo posterior presentado. Si ya existe una presentacion posterior, el estado anterior queda protegido para no romper la continuidad."),
            ("Como veo el detalle, las observaciones y el historial", "Detalle exportable, Observaciones del PLE y Archivos generados aparecen contraidos por defecto. Activa el interruptor de cada seccion para desplegar su informacion."),
            ("Que hago si necesito descargar nuevamente un TXT", "Las descargas son temporales y cada enlace se consume al descargar. Si el boton queda inhabilitado o necesitas otra copia, vuelve a generar el periodo y descarga los archivos nuevos."),
            ("Que control final debo realizar antes de presentar", "Confirma que no existan errores, revisa periodo, CUO, comprobantes y totales, descarga todos los archivos que componen el libro, validalos juntos en el PLE y marca el periodo como presentado solo despues de recibir la constancia.")
        ]);

    private static IReadOnlyCollection<AyudaPreguntaViewModel> CrearPreguntas(
        string prefijo,
        IReadOnlyCollection<(string pregunta, string respuesta)> preguntas)
    {
        var items = new List<AyudaPreguntaViewModel>(preguntas.Count);
        var indice = 1;

        foreach (var (pregunta, respuesta) in preguntas)
        {
            items.Add(new AyudaPreguntaViewModel
            {
                Id = $"{prefijo}-{indice:00}",
                Pregunta = pregunta,
                Respuesta = respuesta
            });

            indice++;
        }

        return items;
    }

    private static string NormalizarClave(string? valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
        {
            return string.Empty;
        }

        return valor.Trim().Replace(" ", string.Empty, StringComparison.Ordinal).ToUpperInvariant();
    }
}
