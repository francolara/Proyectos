using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AsientoFormViewModel
{
    public int? IdAsiento { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione un origen.")]
    public int? IdOrigen { get; set; }

    public string OrigenTexto { get; set; } = string.Empty;

    [DataType(DataType.Date)]
    public DateOnly FechaEmision { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [DataType(DataType.Date)]
    public DateOnly FechaAsiento { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required(ErrorMessage = "Ingrese la glosa del asiento.")]
    [StringLength(500)]
    public string Glosa { get; set; } = string.Empty;

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la moneda.")]
    public int? IdMoneda { get; set; }

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese un tipo de cambio valido.")]
    public decimal TipoCambio { get; set; } = 1m;

    [StringLength(100)]
    public string? ReferenciaExterna { get; set; }

    [StringLength(500)]
    public string? Observacion { get; set; }

    public List<AsientoDetalleFormViewModel> Detalles { get; set; } =
    [
        new()
        {
            Item = 1
        },
        new()
        {
            Item = 2
        }
    ];
}
