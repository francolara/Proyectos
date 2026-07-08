namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IRegistroVentasRepository
{
    Task<IReadOnlyCollection<RegistroVentasItemDto>> ListarAsync(RegistroVentasRequest request, CancellationToken cancellationToken = default);
}
