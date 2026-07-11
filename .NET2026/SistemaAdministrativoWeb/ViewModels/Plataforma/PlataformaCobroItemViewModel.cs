namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class PlataformaCobroItemViewModel
{
    public int IdCuentaAdministradora { get; init; }
    public int IdCuentaAdministradoraSuscripcionPago { get; init; }
    public string NombreCuenta { get; init; } = string.Empty;
    public string CodigoCuenta { get; init; } = string.Empty;
    public string? Contacto { get; init; }
    public string TipoPago { get; init; } = string.Empty;
    public string EstadoPago { get; init; } = string.Empty;
    public decimal Monto { get; init; }
    public string Moneda { get; init; } = "PEN";
    public DateTime FechaPago { get; init; }
    public string? OperacionNumero { get; init; }
    public string? EntidadFinanciera { get; init; }
    public string? ReferenciaExterna { get; init; }
    public string? ProveedorPasarela { get; init; }
    public string? EstadoPasarela { get; init; }
    public string? AccionAplicacion { get; init; }
    public bool AplicarAlConfirmar { get; init; }
    public bool AplicadoSuscripcion { get; init; }
    public string? TipoCobroObjetivo { get; init; }
    public DateOnly? FechaInicioPlanObjetivo { get; init; }
    public string? Observacion { get; init; }
    public DateTime FechaRegistro { get; init; }
}
