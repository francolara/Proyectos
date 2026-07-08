namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AperturaProcesoDetalleDto
{
    public int IdAperturaProcesoDetalle { get; init; }
    public int IdAperturaProceso { get; init; }
    public short Item { get; init; }
    public string TipoDetalle { get; init; } = string.Empty;
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambioAplicado { get; init; }
    public string? TipoDocumento { get; init; }
    public string? Serie { get; init; }
    public string? NumeroDocumento { get; init; }
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public decimal TotalImporteS { get; init; }
    public decimal TotalImporteD { get; init; }
    public string? Observacion { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string? UsuarioRegistro { get; init; }
}
