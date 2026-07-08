namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ILibroElectronicoService
{
    Task<PleConsultaResultadoDto> ConsultarAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, CancellationToken cancellationToken = default);
    Task<PleConsultaResultadoDto> ValidarAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, CancellationToken cancellationToken = default);
    Task<PleGenerationResultDto> GenerarAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, string usuarioGeneracion, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, CancellationToken cancellationToken = default);
    PleDownloadPayload? ObtenerDescarga(string token, bool remover = false);
}
