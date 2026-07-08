namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IAnalisisCuentaRepository
{
    Task<IReadOnlyCollection<AnalisisCuentaItemDto>> ListarAsync(AnalisisCuentaRequest request, CancellationToken cancellationToken = default);
}
