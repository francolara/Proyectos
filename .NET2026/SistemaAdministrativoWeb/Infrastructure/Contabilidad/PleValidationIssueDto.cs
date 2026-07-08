namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleValidationIssueDto
{
    public PleValidationSeverity Severidad { get; init; }
    public string Codigo { get; init; } = string.Empty;
    public string Titulo { get; init; } = string.Empty;
    public string Detalle { get; init; } = string.Empty;
    public string Cuo { get; init; } = string.Empty;
    public int? NumeroAsiento { get; init; }
    public DateOnly? FechaOperacion { get; init; }
    public decimal? TotalDebe { get; init; }
    public decimal? TotalHaber { get; init; }
    public decimal? Diferencia { get; init; }
}
