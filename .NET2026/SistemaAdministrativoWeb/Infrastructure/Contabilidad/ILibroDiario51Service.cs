namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ILibroDiario51Service
{
    Task<IReadOnlyCollection<LibroDiario51Dto>> ListarAsync(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default);
}
