namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class VentaDto : VentaResumenDto
{
    public string TipoDocumentoCliente { get; init; } = string.Empty;
    public string NumeroDocumentoCliente { get; init; } = string.Empty;
    public string DescripcionConfiguracion { get; init; } = string.Empty;
    public List<VentaDetalleDto> Detalles { get; init; } = [];
}
