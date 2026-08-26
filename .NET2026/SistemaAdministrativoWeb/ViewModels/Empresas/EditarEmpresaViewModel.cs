using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Empresas;

public sealed class EditarEmpresaViewModel
{
    [Range(1, int.MaxValue, ErrorMessage = "Seleccione una empresa valida.")]
    public int IdEmpresa { get; set; }

    [Required(ErrorMessage = "Ingrese la razon social.")]
    [StringLength(200, ErrorMessage = "La razon social no puede exceder 200 caracteres.")]
    public string RazonSocial { get; set; } = string.Empty;

    [StringLength(200, ErrorMessage = "El nombre comercial no puede exceder 200 caracteres.")]
    public string? NombreComercial { get; set; }

    [Required(ErrorMessage = "Ingrese el RUC.")]
    [RegularExpression(@"^\d{11}$", ErrorMessage = "El RUC debe contener exactamente 11 digitos.")]
    public string Ruc { get; set; } = string.Empty;
}
