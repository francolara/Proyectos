namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoComprobanteDto
{
    public int IdTipoComprobante { get; init; }
    public string CodigoTipoComprobante { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public bool UsoCompras { get; init; }
    public bool UsoVentas { get; init; }
    public bool Estado { get; init; }
}
