using System.Text.Json;
using GestionServicios.Web.Navigation;

namespace GestionServicios.Web.ViewModels.Security;

// Firma: FRANCO LARA - 14/11/2026 | Ordena los permisos según el menú y muestra los accesos compartidos en su orden original. Identifica permisos con acciones combinadas.
public static class ProfilePermissions
{
    public sealed record Section(string Name, IReadOnlyList<Module> Modules);
    public sealed record Module(string Code, string Name, string[] Permissions, string? Owner = null);

    private static readonly (string Code, string Name, string Section)[] Additional =
    [
        ("CONFIRMACIONES", "Confirmación del cliente", "Seguimiento"),
        ("EJECUCION", "Ejecución del trabajo", "Seguimiento"),
        ("VACACIONES", "Vacaciones", "Planillas"),
        ("ASISTENCIA", "Asistencia", "Planillas"),
        ("COSTOS_PROVEEDOR", "Costos de proveedores por presupuesto", "Logística"),
        ("PERFILES", "Perfiles de acceso", "Administración")
    ];

    public static IReadOnlyList<Section> Sections(IEnumerable<JsonElement> catalog)
    {
        var permissions = catalog.Select(item => item.Text("Codigo", "")).Where(code => code.Length > 0)
            .Distinct(StringComparer.Ordinal).ToLookup(ModuleOf, StringComparer.Ordinal);
        var menu = HytMenu.Items.ToLookup(item => ModuleOf(item.Permission), StringComparer.Ordinal);
        var assigned = new HashSet<string>(StringComparer.Ordinal);
        var sections = new List<Section>();
        foreach (var group in HytMenu.Items.GroupBy(item => item.Group))
        {
            var rows = new List<Module>();
            foreach (var option in group)
            {
                var code = ModuleOf(option.Permission);
                if (!permissions.Contains(code)) continue;
                var primary = menu[code].First();
                rows.Add(assigned.Add(code)
                    ? new(code, option.Label, Ordered(permissions[code]))
                    : new(code, option.Label, [], primary.Group));
            }
            foreach (var entry in Additional.Where(entry => entry.Section == group.Key && permissions.Contains(entry.Code)))
            {
                if (assigned.Add(entry.Code))
                    rows.Add(new(entry.Code, entry.Name, Ordered(permissions[entry.Code])));
            }
            if (rows.Count > 0) sections.Add(new(group.Key, rows));
        }
        var other = permissions.Where(group => !assigned.Contains(group.Key))
            .Select(group => new Module(group.Key, SecurityDisplay.Module(group.Key), Ordered(group)))
            .OrderBy(module => module.Name).ToArray();
        if (other.Length > 0) sections.Add(new("Otros permisos", other));
        return sections;
    }

    public static int Column(string code) => code.Split('.').Last() switch
    {
        "CONSULTAR" => 0,
        "CREAR" => 1,
        "MODIFICAR" or "EDITAR" => 2,
        _ => 3
    };

    public static string? RegisteredAction(string code) => code switch
    {
        "PROVEEDORES.EDITAR" or "TRABAJADORES.EDITAR" or "OBLIGACIONES.EDITAR" or "PERFILES.EDITAR" => "Crear y editar",
        "PAGOS_PROVEEDOR.EDITAR" => "Registrar compras y pagos",
        "FACTURADO.EDITAR" => "Registrar y gestionar comprobantes",
        "FACTORING.EDITAR" => "Registrar y cambiar estado",
        "COSTOS_PROVEEDOR.EDITAR" => "Registrar costos",
        "PAGOS_PERSONAL.EDITAR" => "Guardar planificación y confirmar pago",
        "VACACIONES.EDITAR" => "Registrar y gestionar vacaciones",
        "ASISTENCIA.EDITAR" => "Registrar jornadas y aprobar horas extra",
        _ => null
    };

    private static string ModuleOf(string code) => code.Split('.')[0];
    private static string[] Ordered(IEnumerable<string> permissions) =>
        permissions.OrderBy(Column).ThenBy(SecurityDisplay.Permission).ThenBy(code => code, StringComparer.Ordinal).ToArray();
}
