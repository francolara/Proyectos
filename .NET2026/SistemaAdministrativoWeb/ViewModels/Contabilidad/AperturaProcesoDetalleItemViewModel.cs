namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AperturaProcesoDetalleItemViewModel
{
    public short Item { get; set; }
    public string TipoDetalle { get; set; } = string.Empty;
    public int IdPlanCuenta { get; set; }
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string CodigoMoneda { get; set; } = string.Empty;
    public decimal TipoCambioAplicado { get; set; }
    public string? TipoDocumento { get; set; }
    public string? Serie { get; set; }
    public string? NumeroDocumento { get; set; }
    public decimal Debe { get; set; }
    public decimal Haber { get; set; }
    public decimal TotalImporteS { get; set; }
    public decimal TotalImporteD { get; set; }
    public string? Observacion { get; set; }
}
