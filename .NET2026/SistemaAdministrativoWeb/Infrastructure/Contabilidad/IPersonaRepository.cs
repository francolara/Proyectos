using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPersonaRepository
{
    Task<PagedResult<PersonaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, string? tipoPersona, bool soloClientes, bool soloProveedores, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<PersonaDetalleDto?> ObtenerPorIdAsync(int idEmpresa, int idPersona, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<TipoDocumentoIdentidadDto>> ListarTiposDocumentoAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<UbigeoDepartamentoDto>> ListarDepartamentosAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<UbigeoProvinciaDto>> ListarProvinciasAsync(string codigoDepartamento, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<UbigeoDistritoDto>> ListarDistritosAsync(string codigoProvincia, CancellationToken cancellationToken = default);
    Task<PersonaDetalleDto> GuardarAsync(GuardarPersonaRequest request, CancellationToken cancellationToken = default);
}
