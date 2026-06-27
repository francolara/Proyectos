namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CompraDetraccionPendienteDto
{
    public int IdCompraDetraccion { get; init; }
    public int IdCompra { get; init; }
    public int IdProveedor { get; init; }
    public string CodigoProveedor { get; init; } = string.Empty;
    public string NombreProveedor { get; init; } = string.Empty;
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public DateOnly FechaContabilizacion { get; init; }
    public int IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public string CodigoDetraccionSunat { get; init; } = string.Empty;
    public string DescripcionDetraccion { get; init; } = string.Empty;
    public decimal PorcentajeDetraccion { get; init; }
    public decimal ImporteDetraccion { get; init; }
    public decimal Saldo { get; init; }
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public string ReferenciaDocumento { get; init; } = string.Empty;
    public string Estado { get; init; } = string.Empty;
}
