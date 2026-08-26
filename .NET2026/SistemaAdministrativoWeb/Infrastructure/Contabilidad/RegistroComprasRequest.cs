namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class RegistroComprasRequest
{
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public byte Mes { get; init; }
    public string CodigoPersona { get; init; } = string.Empty;
    public string NumeroComprobante { get; init; } = string.Empty;
}
