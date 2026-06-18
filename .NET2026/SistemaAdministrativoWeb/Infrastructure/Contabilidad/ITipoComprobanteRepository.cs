namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ITipoComprobanteRepository
{
    Task<IReadOnlyCollection<TipoComprobanteDto>> ListarActivosAsync(bool usoCompras, bool usoVentas, CancellationToken cancellationToken = default);
}
