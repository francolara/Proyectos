namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IMigoTipoCambioApiClient
{
    Task<MigoTipoCambioItemDto?> ObtenerPorFechaAsync(DateOnly fecha, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<MigoTipoCambioItemDto>> ObtenerPorRangoAsync(DateOnly fechaInicio, DateOnly fechaFin, CancellationToken cancellationToken = default);
}
