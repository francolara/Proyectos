namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AjusteCuentaProcesoDetalleItemViewModel
{
    public int IdPlanCuenta { get; set; }
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string CodigoMoneda { get; set; } = string.Empty;
    public decimal TipoCambioAplicado { get; set; }
    public int TotalAnalisis { get; set; }
    public int? IdAsiento { get; set; }
    public int? NumeroAsiento { get; set; }
    public decimal TotalDebe { get; set; }
    public decimal TotalHaber { get; set; }
    public string Estado { get; set; } = string.Empty;
    public string? Observacion { get; set; }
}
