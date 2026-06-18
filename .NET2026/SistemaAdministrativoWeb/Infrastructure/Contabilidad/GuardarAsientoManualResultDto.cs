namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarAsientoManualResultDto
{
    public int IdAsiento { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public int NumeroAsiento { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public string Estado { get; init; } = string.Empty;
}
