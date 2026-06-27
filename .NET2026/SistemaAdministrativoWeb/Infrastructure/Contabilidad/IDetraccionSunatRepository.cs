namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IDetraccionSunatRepository
{
    Task<IReadOnlyCollection<DetraccionSunatDto>> ListarActivasAsync(CancellationToken cancellationToken = default);
}
