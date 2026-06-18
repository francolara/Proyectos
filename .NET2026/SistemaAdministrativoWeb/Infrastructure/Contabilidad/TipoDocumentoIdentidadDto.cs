namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoDocumentoIdentidadDto
{
    public string CodigoSunat { get; init; } = string.Empty;
    public string CodigoInterno { get; init; } = string.Empty;
    public string Nombre { get; init; } = string.Empty;
    public byte Orden { get; init; }
}
