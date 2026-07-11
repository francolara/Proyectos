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
    [StringLength(30)]
    public string TipoPago { get; init; } = "TRANSFERENCIA";

    public bool CobroConfirmado { get; init; } = true;
    public DateTime? FechaPago { get; init; }
    public DateOnly? FechaVencimiento { get; init; }

    [StringLength(100)]
    public string? OperacionNumero { get; init; }

    [StringLength(120)]
    public string? EntidadFinanciera { get; init; }

    [StringLength(120)]
    public string? ReferenciaExterna { get; init; }

    [StringLength(50)]
    public string? ProveedorPasarela { get; init; }

    [StringLength(120)]
    public string? TransaccionPasarelaId { get; init; }

    [StringLength(120)]
    public string? PagoPasarelaId { get; init; }

    [StringLength(30)]
    public string? EstadoPasarela { get; init; }

    public string? PayloadPasarela { get; init; }

    [StringLength(500)]
    public string? Observacion { get; init; }

    [StringLength(30)]
    public string? AccionAplicacion { get; init; }

    public bool AplicarAlConfirmar { get; init; }

    [StringLength(20)]
    public string? TipoCobroObjetivo { get; init; }

    public DateOnly? FechaInicioPlanObjetivo { get; init; }

    [Range(0, 60)]
    public int? DiasGraciaObjetivo { get; init; }
}
