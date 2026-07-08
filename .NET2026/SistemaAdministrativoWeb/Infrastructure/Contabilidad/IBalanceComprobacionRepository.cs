namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IBalanceComprobacionRepository
{
    Task<IReadOnlyCollection<BalanceComprobacionItemDto>> ListarAsync(BalanceComprobacionRequest request, CancellationToken cancellationToken = default);
}
