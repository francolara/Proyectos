namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarCompraResultDto
{
    public int IdCompra { get; init; }
    public int? IdAsiento { get; init; }
    public decimal ImporteTotal { get; init; }
    public string Estado { get; init; } = string.Empty;
}
