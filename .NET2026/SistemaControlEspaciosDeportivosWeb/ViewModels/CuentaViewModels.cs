using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class CambiarContrasenaViewModel
{
    public string CorreoUsuario { get; set; } = string.Empty;

    [Required(ErrorMessage = "La contrasena actual es obligatoria.")]
    [DataType(DataType.Password)]
    public string ContrasenaActual { get; set; } = string.Empty;

    [Required(ErrorMessage = "La nueva contrasena es obligatoria.")]
    [StringLength(100, ErrorMessage = "La nueva contrasena debe tener al menos {2} caracteres.", MinimumLength = 6)]
    [DataType(DataType.Password)]
    public string NuevaContrasena { get; set; } = string.Empty;

    [Required(ErrorMessage = "Confirma la nueva contrasena.")]
    [DataType(DataType.Password)]
    [Compare(nameof(NuevaContrasena), ErrorMessage = "La confirmacion no coincide con la nueva contrasena.")]
    public string ConfirmarNuevaContrasena { get; set; } = string.Empty;
}
