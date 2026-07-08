namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CierreProcesoDetalleItemViewModel
{
    public string TipoCierre { get; init; } = string.Empty;
    public string DescripcionCierre { get; init; } = string.Empty;
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambioAplicado { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string? Observacion { get; init; }
}
