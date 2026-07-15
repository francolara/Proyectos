using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PlanCuentaFormViewModel
{
    public int? IdPlanCuenta { get; set; }

    [Display(Name = "Cuenta padre")]
    public int? IdPlanCuentaPadre { get; set; }

    public string CuentaPadreTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el codigo de cuenta.")]
    [StringLength(20, ErrorMessage = "El codigo de cuenta no puede exceder 20 caracteres.")]
    public string CodigoCuenta { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el nombre de la cuenta.")]
    [StringLength(200, ErrorMessage = "El nombre de la cuenta no puede exceder 200 caracteres.")]
    public string NombreCuenta { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese la columna de balance.")]
    [StringLength(1, ErrorMessage = "La columna de balance no puede exceder 1 caracter.")]
    [RegularExpression("^(S|I|N|F|R)$", ErrorMessage = "Seleccione una columna de balance valida.")]
    public string ColBalance { get; set; } = string.Empty;

    [StringLength(3, ErrorMessage = "La moneda no puede exceder 3 caracteres.")]
    [RegularExpression("^(|PEN|USD|S|D)$", ErrorMessage = "Seleccione una moneda valida.")]
    public string? IdMoneda { get; set; }

    [StringLength(1, ErrorMessage = "El tipo de cambio no puede exceder 1 caracter.")]
    [RegularExpression("^(|V|C)$", ErrorMessage = "Seleccione un tipo de cambio valido.")]
    public string? TipoCambio { get; set; }

    public bool AceptaMovimiento { get; set; }
    public bool GeneraDiferenciaPorAnalisis { get; set; }
    public bool RequiereCentroCosto { get; set; }
    public bool Estado { get; set; } = true;
    public bool PermiteConfigurarDestinos { get; set; }
    public PlanCuentaDestinoConfiguracionViewModel ConfiguracionDestino { get; set; } = new();
}
