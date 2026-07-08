namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CierreProcesoDto
{
    public int IdCierreProceso { get; init; }
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public int IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public DateOnly FechaAsiento { get; init; }
    public bool UsaTipoCambioSbs { get; init; }
    public decimal TipoCambioCompra { get; init; }
    public decimal TipoCambioVenta { get; init; }
    public bool ProcesaGananciasPerdidas { get; init; }
    public bool ProcesaInventarios { get; init; }
    public int TotalCuentas { get; init; }
    public int TotalAsientos { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string? UsuarioRegistro { get; init; }
    public List<CierreProcesoDetalleDto> Detalles { get; init; } = [];
}
