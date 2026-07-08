using System.Security.Claims;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public static class LibroElectronicoPermissions
{
    public const string ClaimType = "permiso";
    public const string Ver = "libros-electronicos.ver";
    public const string Consultar = "libros-electronicos.consultar";
    public const string Previsualizar = "libros-electronicos.previsualizar";
    public const string Validar = "libros-electronicos.validar";
    public const string GenerarTxt = "libros-electronicos.generar";
    public const string DescargarTxt = "libros-electronicos.descargar";
    public const string VerHistorial = "libros-electronicos.historial";

    public static bool TienePermiso(ClaimsPrincipal user, string permiso)
    {
        return user.IsInRole("SuperAdmin")
            || user.IsInRole("AdministradorEmpresa")
            || user.Claims.Any(x => string.Equals(x.Type, ClaimType, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.Value, permiso, StringComparison.OrdinalIgnoreCase));
    }
}
