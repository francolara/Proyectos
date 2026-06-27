using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PlanCuentaDestinoConfiguracionViewModel
{
    public bool Activo { get; set; } = true;

    [StringLength(500)]
    public string? Observacion { get; set; }

    public List<CuentaDestinoReglaDetalleFormViewModel> Detalles { get; set; } = [];
}
