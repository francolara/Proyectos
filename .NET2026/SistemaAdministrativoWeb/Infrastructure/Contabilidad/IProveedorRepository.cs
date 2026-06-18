namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IProveedorRepository
{
    Task<IReadOnlyCollection<ProveedorDto>> ListarActivosPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default);
}
