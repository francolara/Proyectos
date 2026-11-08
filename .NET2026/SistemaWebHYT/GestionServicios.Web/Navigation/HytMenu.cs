namespace GestionServicios.Web.Navigation;

public sealed record HytMenuItem(string Code, string Label, string Group, string Permission);
// Firma: FRANCO LARA - 05/11/2026 | Clientes, Tiendas y Contactos como accesos independientes con sus permisos.
public static class HytMenu
{
    public static readonly IReadOnlyList<HytMenuItem> Items =
    [
        new("clientes", "Clientes", "Maestros", "CLIENTES.CONSULTAR"),
        new("tiendas", "Tiendas", "Maestros", "TIENDAS.CONSULTAR"),
        new("contactos", "Contactos", "Maestros", "CLIENTES.CONSULTAR"),
        new("materiales", "Materiales y servicios", "Maestros", "MATERIALES.CONSULTAR"),
        new("especialidades", "Especialidades", "Maestros", "MATERIALES.CONSULTAR"),
        new("tipos-servicio", "Tipos de servicio", "Maestros", "MATERIALES.CONSULTAR"),
        new("unidades", "Unidades de medida", "Maestros", "UNIDADES.CONSULTAR"),
        new("zonas", "Zonas", "Maestros", "ZONAS.CONSULTAR"),
        new("tarifarios", "Tarifarios", "Comercial", "TARIFARIOS.CONSULTAR"),
        new("presupuestos", "Presupuestos", "Comercial", "PRESUPUESTOS.CONSULTAR"),
        new("oc-cliente", "Orden de Compra - Cliente", "Seguimiento", "OC_CLIENTE.CONSULTAR"),
        new("guias", "Guía", "Seguimiento", "GUIAS.CONSULTAR"),
        new("reportes", "Reportes", "Consultas", "REPORTES.CONSULTAR"),
        new("usuarios", "Usuarios y permisos", "Administración", "USUARIOS.CONSULTAR"),
        new("configuracion", "Configuración HYT", "Administración", "CONFIGURACION.CONSULTAR"),
        new("auditoria", "Auditoría", "Administración", "AUDITORIA.CONSULTAR")
    ];
}
