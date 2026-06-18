namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoPreviewResultDto
{
    public string ModuloOperacion { get; init; } = string.Empty;
    public int IdConfiguracionContabilizacion { get; init; }
    public string EscenarioOperacion { get; init; } = string.Empty;
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string GlosaAsiento { get; init; } = string.Empty;
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public bool Cuadrado { get; init; }
    public string? MensajeValidacion { get; init; }
    public List<AsientoPreviewLineaDto> Lineas { get; init; } = [];
}
