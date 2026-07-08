namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CompraDto : CompraResumenDto
{
    public string TipoDocumentoProveedor { get; init; } = string.Empty;
    public string NumeroDocumentoProveedor { get; init; } = string.Empty;
    public string DescripcionConfiguracion { get; init; } = string.Empty;
    public int? IdCompraRetencion { get; init; }
    public int? IdCompraDetraccion { get; init; }
    public int? IdAsientoDetraccion { get; init; }
    public int? IdCompraPercepcion { get; init; }
    public int? IdAsientoPercepcion { get; init; }
    public string CodigoDetraccionSunat { get; init; } = string.Empty;
    public string DescripcionDetraccionSunat { get; init; } = string.Empty;
    public string CodigoPercepcion { get; init; } = string.Empty;
    public string DescripcionPercepcion { get; init; } = string.Empty;
    public List<CompraDetalleDto> Detalles { get; init; } = [];
}
