namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleResumenDto
{
    public string Empresa { get; init; } = string.Empty;
    public string Ruc { get; init; } = string.Empty;
    public string Libro { get; init; } = string.Empty;
    public string Formato { get; init; } = string.Empty;
    public string Periodo { get; init; } = string.Empty;
    public int CantidadAsientos { get; init; }
    public int CantidadMovimientos { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public string NombreArchivo { get; init; } = string.Empty;
}
