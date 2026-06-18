namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IMonedaRepository
{
    Task<IReadOnlyCollection<MonedaDto>> ListarActivasAsync(CancellationToken cancellationToken = default);
}
