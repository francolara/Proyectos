namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoCambioPeriodoSyncResult
{
    public int TotalConsultados { get; init; }
    public int TotalSincronizados { get; init; }
    public IReadOnlyCollection<TipoCambioDto> TiposCambio { get; init; } = Array.Empty<TipoCambioDto>();
}
