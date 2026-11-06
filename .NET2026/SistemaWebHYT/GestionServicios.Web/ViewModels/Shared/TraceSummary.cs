using System.Globalization;
namespace GestionServicios.Web.ViewModels.Shared;

// Firma: FRANCO LARA - 05/11/2026 | Fechas y responsables compartidos con formato Perú y ausencia de actualizaciones.
public sealed record TraceSummary(string? CreatedBy, string CreatedAt, string? UpdatedBy, string UpdatedAt)
{
    public static TraceSummary From(Func<string,string> text)
    {
        string Date(string key,string fallback)=>DateTimeOffset.TryParse(text(key),CultureInfo.InvariantCulture,DateTimeStyles.AssumeUniversal,out var date)?PeruDate.Display(date):fallback;
        return new(text("UsuarioRegistro"),Date("FechaRegistro","—"),text("UsuarioActualizacion"),Date("FechaActualizacion","Sin actualizaciones"));
    }
}
