using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleConsultaResultadoDto
{
    public string LibroElectronico { get; init; } = PleLibroElectronicoCatalogo.LibroDiario51;
    public PleResumenDto Resumen { get; init; } = new();
    public PleValidationResultDto Validacion { get; init; } = new();
    public IReadOnlyCollection<LibroDiario51Dto> LibroDiario51Items { get; init; } = [];
    public IReadOnlyCollection<LibroDiario52Dto> LibroDiario52Items { get; init; } = [];
    public IReadOnlyCollection<LibroMayor61Dto> LibroMayor61Items { get; init; } = [];
    public int TotalRegistrosPreview { get; init; }
    public int PaginaPreview { get; init; }
    public int TamanoPaginaPreview { get; init; }
    public PagedResult<PleHistoryItemDto> Historial { get; init; } = new();
    public PlePresentacionContextoDto Presentacion { get; init; } = new();
    public bool GeneracionBloqueada { get; init; }
    public string MensajeBloqueoGeneracion { get; init; } = string.Empty;
}
