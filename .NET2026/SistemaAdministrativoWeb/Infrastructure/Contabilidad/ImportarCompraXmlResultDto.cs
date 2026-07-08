namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ImportarCompraXmlResultDto
{
    public int IdCompra { get; init; }
    public string Estado { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
}
