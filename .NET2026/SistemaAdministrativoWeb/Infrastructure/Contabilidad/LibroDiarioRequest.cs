namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroDiarioRequest
{
    public int IdEmpresa { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public string Moneda { get; init; } = "PEN";
    public string Modo { get; init; } = "A";
    public string? CuentaDesde { get; init; }
    public string? CuentaHasta { get; init; }
    public string? OrigenDesde { get; init; }
    public string? OrigenHasta { get; init; }
}
