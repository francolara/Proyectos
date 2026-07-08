namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AperturaProcesoDto
{
    public int IdAperturaProceso { get; init; }
    public int IdEmpresa { get; init; }
    public short AnioApertura { get; init; }
    public short AnioSaldo { get; init; }
    public byte MesSaldoHasta { get; init; }
    public string PeriodoSaldoHasta { get; init; } = string.Empty;
    public int IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public DateOnly FechaAsiento { get; init; }
    public bool UsaTipoCambioSbs { get; init; }
    public decimal TipoCambioCompra { get; init; }
    public decimal TipoCambioVenta { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public int TotalLineas { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string? UsuarioRegistro { get; init; }
    public List<AperturaProcesoDetalleDto> Detalles { get; init; } = [];
}
