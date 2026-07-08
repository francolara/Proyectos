namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroMayor61Dto
{
    public string PeriodoPle { get; init; } = string.Empty;
    public string Cuo { get; init; } = string.Empty;
    public string CorrelativoMovimiento { get; init; } = string.Empty;
    public string CodigoCuentaContable { get; init; } = string.Empty;
    public DateOnly FechaOperacion { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public string EstadoOperacion { get; init; } = PleEstadoRegistroCatalogo.Vigente;
    public int NumeroAsiento { get; init; }
}
