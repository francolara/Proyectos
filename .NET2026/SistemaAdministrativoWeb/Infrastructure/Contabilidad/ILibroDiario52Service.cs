namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ILibroDiario52Service
{
    Task<IReadOnlyCollection<LibroDiario52Dto>> ListarAsync(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default);
}
