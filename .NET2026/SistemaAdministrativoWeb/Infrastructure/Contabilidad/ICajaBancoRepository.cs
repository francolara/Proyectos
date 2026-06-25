using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ICajaBancoRepository
{
    Task<PagedResult<CajaBancoResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, int? idBancoConfiguracionEmpresa, short anio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<CajaBancoResumenCuentaDto> ObtenerResumenCuentaAsync(int idEmpresa, int? idBancoConfiguracionEmpresa, short anio, byte mes, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<OperacionBancariaDto>> ListarOperacionesBancariasAsync(string tipoMovimiento, string? textoBusqueda = null, int tamanoPagina = 100, string? idTipoOpeBancaria = null, CancellationToken cancellationToken = default);
    Task<CajaBancoMovimientoDto?> ObtenerPorIdAsync(int idEmpresa, int idMovimientoBanco, CancellationToken cancellationToken = default);
    Task<CajaBancoMovimientoDto> GuardarAsync(GuardarCajaBancoRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idEmpresa, int idMovimientoBanco, CancellationToken cancellationToken = default);
    Task<PagedResult<TransferenciaCuentaResumenDto>> ListarTransferenciasPaginadoPorEmpresaAsync(int idEmpresa, short anio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<TransferenciaCuentaResultadoDto> GuardarTransferenciaAsync(GuardarTransferenciaCuentaRequest request, CancellationToken cancellationToken = default);
    Task EliminarTransferenciaAsync(int idEmpresa, int idMovimientoBancoEmisor, CancellationToken cancellationToken = default);
}
