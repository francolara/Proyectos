namespace SistemaAdministrativoWeb.Infrastructure.Empresas;

public interface IEmpresaRepository
{
    Task<IReadOnlyCollection<EmpresaDisponibleDto>> ListarPorUsuarioAsync(string aspNetUserId, CancellationToken cancellationToken = default);
}
