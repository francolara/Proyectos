using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionContabilizacionDetalleFormViewModel
{
    [Range(1, short.MaxValue, ErrorMessage = "El orden debe ser mayor a cero.")]
    public short Orden { get; set; }

    [Required(ErrorMessage = "Seleccione el componente contable.")]
    [StringLength(20, ErrorMessage = "El componente contable no puede exceder 20 caracteres.")]
    public string ComponenteContable { get; set; } = "BRUTO";

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione una cuenta.")]
    public int? IdPlanCuenta { get; set; }

    public string CuentaTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Seleccione la naturaleza.")]
    [StringLength(1, ErrorMessage = "La naturaleza no puede exceder 1 caracter.")]
    public string NaturalezaMovimiento { get; set; } = "D";

    public bool Activo { get; set; } = true;
}
