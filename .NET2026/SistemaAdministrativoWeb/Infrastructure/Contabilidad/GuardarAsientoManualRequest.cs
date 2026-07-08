namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarAsientoManualRequest
{
    public int? IdAsiento { get; init; }
    public int IdEmpresa { get; init; }
    public int IdOrigen { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public DateOnly FechaAsiento { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public int IdMoneda { get; init; }
    public decimal TipoCambio { get; init; }
    public string? ReferenciaExterna { get; init; }
    public string? Observacion { get; init; }
    public string? UsuarioRegistro { get; init; }
    public IReadOnlyCollection<GuardarAsientoDetalleRequest> Detalles { get; init; } = [];
}
