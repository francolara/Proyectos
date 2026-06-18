using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaDestinoReglaFormViewModel
{
    [Range(2000, 9999, ErrorMessage = "Ingrese un ejercicio valido.")]
    public short Ejercicio { get; set; } = (short)DateTime.Today.Year;

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la cuenta origen.")]
    public int? IdPlanCuentaOrigen { get; set; }

    [StringLength(500)]
    public string? Observacion { get; set; }

    public bool Activo { get; set; } = true;

    public List<CuentaDestinoReglaDetalleFormViewModel> Detalles { get; set; } =
    [
        new()
        {
            Orden = 1,
            Porcentaje = 100,
            Activo = true
        }
    ];
}
