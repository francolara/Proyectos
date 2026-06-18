using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaDestinoReglaDetalleFormViewModel
{
    public int? IdCuentaDestinoReglaDetalle { get; set; }

    [Range(1, short.MaxValue, ErrorMessage = "El orden debe ser mayor a cero.")]
    public short Orden { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la cuenta destino cargo.")]
    public int? IdPlanCuentaDestinoCargo { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la cuenta destino abono.")]
    public int? IdPlanCuentaDestinoAbono { get; set; }

    [Range(typeof(decimal), "0.0001", "100", ErrorMessage = "El porcentaje debe estar entre 0 y 100.")]
    public decimal Porcentaje { get; set; }

    public bool Activo { get; set; } = true;
}
