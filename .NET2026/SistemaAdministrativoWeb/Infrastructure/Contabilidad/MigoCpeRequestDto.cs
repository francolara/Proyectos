using System.Globalization;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MigoCpeRequestDto
{
    public string RucEmisor { get; init; } = string.Empty;
    public string TipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public decimal? Monto { get; init; }

    public string FechaEmisionTexto => FechaEmision.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
    public string? MontoTexto => Monto.HasValue ? Monto.Value.ToString("0.##", CultureInfo.InvariantCulture) : null;
}
