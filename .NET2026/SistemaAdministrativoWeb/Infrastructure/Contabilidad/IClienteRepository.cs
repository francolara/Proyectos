namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IClienteRepository
{
    Task<IReadOnlyCollection<ClienteDto>> ListarActivosPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default);
}
