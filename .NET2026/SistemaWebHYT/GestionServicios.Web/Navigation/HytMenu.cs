namespace GestionServicios.Web.Navigation;

public sealed record HytMenuItem(string Code, string Label, string Group, string Permission);
// Firma: FRANCO LARA - 05/11/2026 | Clientes, Tiendas y Contactos como accesos independientes con sus permisos.
// Firma: FRANCO LARA - 10/11/2026 | Ordena los accesos de Maestros y Planillas y renombra Trabajos de Proveedores y agrega la bandeja de pagos y los reportes por tipo.
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
        new("reportes", "Reportes", "Consultas", "REPORTES.CONSULTAR"),
        // Firma: FRANCO LARA - 09/11/2026 | Proveedores, pagos, trabajadores, reporte mensual y obligaciones.
        new("cargos", "Cargos", "Planillas", "TRABAJADORES.CONSULTAR"),
        new("trabajadores", "Trabajadores", "Planillas", "TRABAJADORES.CONSULTAR"),
        new("sctr", "SCTR", "Planillas", "OBLIGACIONES.CONSULTAR"),
        new("pagos-proveedores", "Trabajos de Proveedores", "Planillas", "PAGOS_PROVEEDOR.CONSULTAR"),
        new("pagos-pendientes-proveedores", "Pagos a Proveedores", "Planillas", "PAGOS_PROVEEDOR.CONSULTAR"),
        new("planificacion-pagos", "Planificación de Planilla", "Planillas", "PAGOS_PERSONAL.CONSULTAR"),
        new("obligaciones-planillas", "Obligaciones", "Planillas", "OBLIGACIONES.CONSULTAR"),
        new("reporte-planillas", "Reportes de Planillas", "Planillas", "REPORTES_PLANILLAS.CONSULTAR"),
        new("usuarios", "Usuarios y permisos", "Administración", "USUARIOS.CONSULTAR"),
        new("configuracion", "Configuración HYT", "Administración", "CONFIGURACION.CONSULTAR"),
        new("auditoria", "Auditoría", "Administración", "AUDITORIA.CONSULTAR")
    ];
}
