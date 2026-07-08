namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroMayorRequest
{
    public int IdEmpresa { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public string? CuentaDesde { get; init; }
    public string? CuentaHasta { get; init; }
    public string? NumeroDocumento { get; init; }
}
