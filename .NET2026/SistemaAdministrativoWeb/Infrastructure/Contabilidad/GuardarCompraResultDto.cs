namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarCompraResultDto
{
    public int IdCompra { get; init; }
    public int? IdAsiento { get; init; }
    public int? IdCompraRetencion { get; init; }
    public int? IdAsientoDetraccion { get; init; }
    public int? IdAsientoPercepcion { get; init; }
    public decimal ImporteTotal { get; init; }
    public decimal Retencion { get; init; }
    public decimal ImporteDetraccion { get; init; }
    public decimal ImportePercepcion { get; init; }
    public string Estado { get; init; } = string.Empty;
}
