namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public class CompraResumenDto
{
    public int IdCompra { get; init; }
    public int IdEmpresa { get; init; }
    public int IdProveedor { get; init; }
    public string CodigoProveedor { get; init; } = string.Empty;
    public string NombreProveedor { get; init; } = string.Empty;
    public int IdConfiguracionContabilizacion { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public string EscenarioOperacion { get; init; } = string.Empty;
    public int? IdAsiento { get; init; }
    public DateOnly FechaEmision { get; init; }
    public DateOnly FechaContabilizacion { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public string TipoComprobante { get; init; } = string.Empty;
    public string DescripcionTipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public int IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
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
    public decimal Saldo { get; init; }
    public string? Observacion { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string Situacion { get; init; } = string.Empty;
}
