namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CierreProcesoDetalleItemViewModel
{
    public short Item { get; init; }
    public string TipoCierre { get; init; } = string.Empty;
    public string DescripcionCierre { get; init; } = string.Empty;
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambioAplicado { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public string DH { get; init; } = string.Empty;
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public decimal TotalImporteS { get; init; }
    public decimal TotalImporteD { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string? Observacion { get; init; }
}
