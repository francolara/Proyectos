namespace GestionServicios.Web.Navigation;

public sealed record HytMenuItem(string Code, string Label, string Group, string Permission);
// Firma: FRANCO LARA - 05/11/2026 | Clientes, Tiendas y Contactos como accesos independientes con sus permisos.
// Firma: FRANCO LARA - 10/11/2026 | Ordena los accesos de Maestros y Planillas y renombra Trabajos de Proveedores y agrega la bandeja de pagos y los reportes por tipo, Registro de compras, Emisión de comprobantes y Negociación de Factoring en Administración. Agrupa trabajos y pagos de proveedores en Logística. Agrega Pago de detracciones en Administración y Conceptos de planilla en Planillas.
// Firma: FRANCO LARA - 11/11/2026 | Ubica SCTR en el grupo Seguimiento. Coloca Pago de detracciones después de Registro de compras.
public static class HytMenu
{
    public static readonly IReadOnlyList<HytMenuItem> Items =
    [
        new("proveedores", "Proveedores", "Maestros", "PROVEEDORES.CONSULTAR"),
        new("clientes", "Clientes", "Maestros", "CLIENTES.CONSULTAR"),
        new("tiendas", "Tiendas", "Maestros", "TIENDAS.CONSULTAR"),
        new("contactos", "Contactos", "Maestros", "CLIENTES.CONSULTAR"),
        new("zonas", "Zonas", "Maestros", "ZONAS.CONSULTAR"),
        new("especialidades", "Especialidades", "Maestros", "MATERIALES.CONSULTAR"),
        new("tipos-servicio", "Tipos de servicio", "Maestros", "MATERIALES.CONSULTAR"),
        new("unidades", "Unidades de medida", "Maestros", "UNIDADES.CONSULTAR"),
        new("materiales", "Materiales y servicios", "Maestros", "MATERIALES.CONSULTAR"),
        new("tarifarios", "Tarifarios", "Comercial", "TARIFARIOS.CONSULTAR"),
        new("presupuestos", "Presupuestos", "Comercial", "PRESUPUESTOS.CONSULTAR"),
        new("oc-cliente", "Orden de Compra - Cliente", "Seguimiento", "OC_CLIENTE.CONSULTAR"),
        new("guias", "Guía", "Seguimiento", "GUIAS.CONSULTAR"),
        new("sctr", "SCTR", "Seguimiento", "OBLIGACIONES.CONSULTAR"),
        new("reportes", "Reportes", "Consultas", "REPORTES.CONSULTAR"),
        // Firma: FRANCO LARA - 09/11/2026 | Proveedores, pagos, trabajadores, reporte mensual y obligaciones.
        new("conceptos-planilla", "Conceptos de planilla", "Planillas", "TRABAJADORES.CONSULTAR"),
        new("cargos", "Cargos", "Planillas", "TRABAJADORES.CONSULTAR"),
        new("trabajadores", "Trabajadores", "Planillas", "TRABAJADORES.CONSULTAR"),
        new("planificacion-pagos", "Planificación de Planilla", "Planillas", "PAGOS_PERSONAL.CONSULTAR"),
        new("obligaciones-planillas", "Obligaciones", "Planillas", "OBLIGACIONES.CONSULTAR"),
        new("reporte-planillas", "Reportes de Planillas", "Planillas", "REPORTES_PLANILLAS.CONSULTAR"),
        new("pagos-proveedores", "Trabajos de Proveedores", "Logística", "PAGOS_PROVEEDOR.CONSULTAR"),
        new("pagos-pendientes-proveedores", "Pagos a Proveedores", "Logística", "PAGOS_PROVEEDOR.CONSULTAR"),
        new("provisiones-compras", "Registro de compras", "Administración", "PAGOS_PROVEEDOR.CONSULTAR"),
        new("pago-detracciones", "Pago de detracciones", "Administración", "PAGOS_PROVEEDOR.CONSULTAR"),
        new("emision-comprobantes", "Emisión de comprobantes", "Administración", "FACTURADO.CONSULTAR"),
        new("factoring", "Negociación de Factoring", "Administración", "FACTORING.CONSULTAR"),
        new("usuarios", "Usuarios y permisos", "Administración", "USUARIOS.CONSULTAR"),
        new("configuracion", "Configuración HYT", "Administración", "CONFIGURACION.CONSULTAR"),
        new("auditoria", "Auditoría", "Administración", "AUDITORIA.CONSULTAR")
    ];
}
