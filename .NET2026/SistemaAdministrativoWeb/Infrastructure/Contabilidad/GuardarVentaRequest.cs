namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarVentaRequest
{
    public int? IdVenta { get; init; }
    public int IdEmpresa { get; init; }
    public int IdCliente { get; init; }
    public int IdConfiguracionContabilizacion { get; init; }
    public DateOnly FechaEmision { get; init; }
    public DateOnly FechaContabilizacion { get; init; }
    public string TipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public int IdMoneda { get; init; }
    public decimal TipoCambio { get; init; }
    public decimal BaseImponible { get; init; }
    public decimal TotalExonerado { get; init; }
    public decimal TotalInafecto { get; init; }
    public decimal Icbper { get; init; }
    public decimal Igv { get; init; }
    public decimal Isc { get; init; }
    public decimal OtrosTributos { get; init; }
    public decimal Redondeo { get; init; }
    public decimal ImporteTotal { get; init; }
    public string? Observacion { get; init; }
    public string? UsuarioRegistro { get; init; }
    public IReadOnlyCollection<GuardarVentaDetalleRequest> Detalles { get; init; } = [];
}
