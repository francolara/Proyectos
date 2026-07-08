namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ImportarVentaXmlResultDto
{
    public int IdVenta { get; init; }
    public string Estado { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
}
