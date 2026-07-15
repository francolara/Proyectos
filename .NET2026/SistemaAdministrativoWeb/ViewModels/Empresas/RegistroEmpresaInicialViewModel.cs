using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Empresas;

public sealed class RegistroEmpresaInicialViewModel
{
    public bool EsEmpresaInicial { get; set; } = true;

    [Required(ErrorMessage = "Ingrese el nombre del contacto principal.")]
    [StringLength(180, ErrorMessage = "El nombre del contacto principal no puede exceder 180 caracteres.")]
    public string NombreContacto { get; set; } = string.Empty;

    [StringLength(30, ErrorMessage = "El telefono no puede exceder 30 caracteres.")]
    public string? Telefono { get; set; }

    [Required(ErrorMessage = "Ingrese el correo principal.")]
    [EmailAddress(ErrorMessage = "Ingrese un correo electronico valido.")]
    public string Correo { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese la razon social.")]
    [StringLength(200, ErrorMessage = "La razon social no puede exceder 200 caracteres.")]
    public string RazonSocial { get; set; } = string.Empty;

    [StringLength(200, ErrorMessage = "El nombre comercial no puede exceder 200 caracteres.")]
    public string? NombreComercial { get; set; }

    [Required(ErrorMessage = "Ingrese el RUC.")]
    [StringLength(11, MinimumLength = 11, ErrorMessage = "El RUC debe tener 11 digitos.")]
    public string Ruc { get; set; } = string.Empty;
}
