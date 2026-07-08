namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ImportarCompraXmlRequest
{
    public int IdEmpresa { get; init; }
    public int IdProveedor { get; init; }
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
    public bool ExoneracionRenta4ta { get; init; }
    public decimal PorcentajeRetencion { get; init; }
    public decimal Retencion { get; init; }
    public bool TieneDetraccion { get; init; }
    public int? IdDetraccionSunat { get; init; }
    public decimal PorcentajeDetraccion { get; init; }
    public decimal ImporteDetraccion { get; init; }
    public bool TienePercepcion { get; init; }
    public int? IdTipoPercepcion { get; init; }
    public decimal PorcentajePercepcion { get; init; }
    public decimal BasePercepcion { get; init; }
    public decimal ImportePercepcion { get; init; }
    public string? Observacion { get; init; }
    public string? UsuarioRegistro { get; init; }
    public IReadOnlyCollection<ImportarCompraXmlDetalleRequest> Detalles { get; init; } = [];
}
