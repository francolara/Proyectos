namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CompraDto : CompraResumenDto
{
    public string TipoDocumentoProveedor { get; init; } = string.Empty;
    public string NumeroDocumentoProveedor { get; init; } = string.Empty;
    public string DescripcionConfiguracion { get; init; } = string.Empty;
    public List<CompraDetalleDto> Detalles { get; init; } = [];
}
