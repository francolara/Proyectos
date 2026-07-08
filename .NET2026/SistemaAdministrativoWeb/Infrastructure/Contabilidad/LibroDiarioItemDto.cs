namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroDiarioItemDto
{
    public string Modo { get; init; } = string.Empty;
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string Periodo { get; init; } = string.Empty;
    public int NumeroAsiento { get; init; }
    public short Item { get; init; }
    public DateOnly? FechaEmision { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string NombreAuxiliar { get; init; } = string.Empty;
    public string TipoDocumento { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Referencia { get; init; } = string.Empty;
    public string Glosa { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public decimal DebeDolares { get; init; }
    public decimal HaberDolares { get; init; }
}
