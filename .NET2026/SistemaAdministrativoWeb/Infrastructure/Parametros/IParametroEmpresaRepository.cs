using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Parametros;

public interface IParametroEmpresaRepository
{
    Task<PagedResult<ParametroEmpresaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? tipoParametro, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<ParametroEmpresaDto?> ObtenerAsync(int idEmpresa, int idParametroEmpresa, CancellationToken cancellationToken = default);
    Task<ParametroEmpresaDto> GuardarAsync(GuardarParametroEmpresaRequest request, CancellationToken cancellationToken = default);
    Task CargarDefaultAsync(int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
