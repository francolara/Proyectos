namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AjusteCuentaProcesoDetalleDto
{
    public int IdAjusteCuentaProcesoDetalle { get; init; }
    public int IdAjusteCuentaProceso { get; init; }
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambioAplicado { get; init; }
    public int TotalAnalisis { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string? Observacion { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string? UsuarioRegistro { get; init; }
}
