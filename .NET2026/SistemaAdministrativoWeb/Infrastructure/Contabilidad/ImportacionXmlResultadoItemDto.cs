namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ImportacionXmlResultadoItemDto
{
    public string NombreArchivo { get; init; } = string.Empty;
    public bool Importado { get; init; }
    public string Mensaje { get; init; } = string.Empty;
    public string TipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public string NombreTercero { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
    public int? IdRegistro { get; init; }
}
