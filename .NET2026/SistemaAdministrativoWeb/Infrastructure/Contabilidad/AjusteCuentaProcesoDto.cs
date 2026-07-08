namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AjusteCuentaProcesoDto
{
    public int IdAjusteCuentaProceso { get; init; }
    public int IdEmpresa { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public int IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public DateOnly FechaAsiento { get; init; }
    public int TotalCuentas { get; init; }
    public int TotalAsientos { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string? UsuarioRegistro { get; init; }
    public List<AjusteCuentaProcesoDetalleDto> Detalles { get; init; } = [];
}
