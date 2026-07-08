using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PlanCuentaFormViewModel
{
    public int? IdPlanCuenta { get; set; }

    [Display(Name = "Cuenta padre")]
    public int? IdPlanCuentaPadre { get; set; }

    public string CuentaPadreTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el codigo de cuenta.")]
    [StringLength(20)]
    public string CodigoCuenta { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el nombre de la cuenta.")]
    [StringLength(200)]
    public string NombreCuenta { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese la columna de balance.")]
    [StringLength(1)]
    [RegularExpression("^(S|I|N|F|R)$", ErrorMessage = "Seleccione una columna de balance valida.")]
    public string ColBalance { get; set; } = string.Empty;

    [StringLength(3)]
    [RegularExpression("^(|PEN|USD|S|D)$", ErrorMessage = "Seleccione una moneda valida.")]
    public string? IdMoneda { get; set; }

    [StringLength(1)]
    [RegularExpression("^(|V|C)$", ErrorMessage = "Seleccione un tipo de cambio valido.")]
    public string? TipoCambio { get; set; }

    public bool AceptaMovimiento { get; set; }
    public bool GeneraDiferenciaPorAnalisis { get; set; }
    public bool RequiereCentroCosto { get; set; }
    public bool Estado { get; set; } = true;
    public bool PermiteConfigurarDestinos { get; set; }
    public PlanCuentaDestinoConfiguracionViewModel ConfiguracionDestino { get; set; } = new();
}
