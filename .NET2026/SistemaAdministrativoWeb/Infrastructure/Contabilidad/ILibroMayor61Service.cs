namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ILibroMayor61Service
{
    Task<IReadOnlyCollection<LibroMayor61Dto>> ListarAsync(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default);
}
