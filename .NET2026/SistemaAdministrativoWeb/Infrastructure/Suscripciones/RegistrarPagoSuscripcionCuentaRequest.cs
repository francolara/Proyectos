namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class RegistrarPagoSuscripcionCuentaRequest
{
    public int IdCuentaAdministradora { get; init; }
    public string TipoPago { get; init; } = "TRANSFERENCIA";
    public string EstadoPago { get; init; } = "PAGADO";
    public decimal Monto { get; init; }
    public string Moneda { get; init; } = "PEN";
    public DateTime FechaPago { get; init; }
    public DateOnly? FechaVencimiento { get; init; }
    public string? OperacionNumero { get; init; }
    public string? EntidadFinanciera { get; init; }
    public string? ReferenciaExterna { get; init; }
    public string? ProveedorPasarela { get; init; }
    public string? TransaccionPasarelaId { get; init; }
    public string? PagoPasarelaId { get; init; }
    public string? EstadoPasarela { get; init; }
    public string? PayloadPasarela { get; init; }
    public string? Observacion { get; init; }
    public string? AccionAplicacion { get; init; }
    public bool AplicarAlConfirmar { get; init; }
    public string? TipoCobroObjetivo { get; init; }
    public DateOnly? FechaInicioPlanObjetivo { get; init; }
    public int? DiasGraciaObjetivo { get; init; }
    public string? UsuarioRegistro { get; init; }
}
