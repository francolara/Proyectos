using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class RegistrarPagoSuscripcionCuentaViewModel
{
    [Required]
    public int IdCuentaAdministradora { get; init; }

    [Required]
    [Range(typeof(decimal), "0.01", "999999999")]
    public decimal Monto { get; init; }

    [Required]
    [StringLength(30, ErrorMessage = "El tipo de pago no puede exceder 30 caracteres.")]
    public string TipoPago { get; init; } = "TRANSFERENCIA";

    public bool CobroConfirmado { get; init; } = true;
    public DateTime? FechaPago { get; init; }
    public DateOnly? FechaVencimiento { get; init; }

    [StringLength(100, ErrorMessage = "El numero de operacion no puede exceder 100 caracteres.")]
    public string? OperacionNumero { get; init; }

    [StringLength(120, ErrorMessage = "La entidad financiera no puede exceder 120 caracteres.")]
    public string? EntidadFinanciera { get; init; }

    [StringLength(120, ErrorMessage = "La referencia externa no puede exceder 120 caracteres.")]
    public string? ReferenciaExterna { get; init; }

    [StringLength(50, ErrorMessage = "El proveedor de pasarela no puede exceder 50 caracteres.")]
    public string? ProveedorPasarela { get; init; }

    [StringLength(120, ErrorMessage = "El identificador de transaccion no puede exceder 120 caracteres.")]
    public string? TransaccionPasarelaId { get; init; }

    [StringLength(120, ErrorMessage = "El identificador de pago no puede exceder 120 caracteres.")]
    public string? PagoPasarelaId { get; init; }

    [StringLength(30, ErrorMessage = "El estado de pasarela no puede exceder 30 caracteres.")]
    public string? EstadoPasarela { get; init; }

    public string? PayloadPasarela { get; init; }

    [StringLength(500, ErrorMessage = "La observacion no puede exceder 500 caracteres.")]
    public string? Observacion { get; init; }

    [StringLength(30, ErrorMessage = "La accion de aplicacion no puede exceder 30 caracteres.")]
    public string? AccionAplicacion { get; init; }

    public bool AplicarAlConfirmar { get; init; }

    [StringLength(20, ErrorMessage = "El tipo de cobro objetivo no puede exceder 20 caracteres.")]
    public string? TipoCobroObjetivo { get; init; }

    public DateOnly? FechaInicioPlanObjetivo { get; init; }

    [Range(0, 60)]
    public int? DiasGraciaObjetivo { get; init; }
}
