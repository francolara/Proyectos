namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarVentaResultDto
{
    public int IdVenta { get; init; }
    public int? IdAsiento { get; init; }
    public decimal ImporteTotal { get; init; }
    public string Estado { get; init; } = string.Empty;
}
