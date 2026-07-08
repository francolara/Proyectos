namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IRegistroComprasRepository
{
    Task<IReadOnlyCollection<RegistroComprasItemDto>> ListarAsync(RegistroComprasRequest request, CancellationToken cancellationToken = default);
}
