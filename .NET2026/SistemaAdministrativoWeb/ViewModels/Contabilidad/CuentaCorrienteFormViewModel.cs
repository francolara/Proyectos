using System.ComponentModel.DataAnnotations;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaCorrienteFormViewModel
{
    public int? IdBancoConfiguracionEmpresa { get; set; }

    [Required(ErrorMessage = "Seleccione un banco.")]
    public int? IdBanco { get; set; }

    public string BancoTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el numero de cuenta corriente.")]
    [StringLength(50, ErrorMessage = "El numero de cuenta no puede exceder 50 caracteres.")]
    public string NroCuentaCorriente { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el titular de la cuenta.")]
    [StringLength(200, ErrorMessage = "El titular no puede exceder 200 caracteres.")]
    public string Titular { get; set; } = string.Empty;

    [Required(ErrorMessage = "Seleccione la moneda.")]
    public int? IdMoneda { get; set; }

    [Required(ErrorMessage = "Seleccione la cuenta contable asociada.")]
    public int? IdPlanCuenta { get; set; }

    public string CuentaTexto { get; set; } = string.Empty;

    public bool Activo { get; set; } = true;

    public DateTime? FechaRegistro { get; set; }
    public string UsuarioRegistro { get; set; } = string.Empty;

    public List<MonedaDto> Monedas { get; set; } = [];
}
