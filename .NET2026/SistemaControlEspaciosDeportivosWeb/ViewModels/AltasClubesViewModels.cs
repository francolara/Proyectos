using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class AltaClubSolicitudFormViewModel
{
    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string NombreContacto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(30, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Telefono { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [EmailAddress(ErrorMessage = "Ingresa un correo electronico valido.")]
    public string Correo { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [DataType(DataType.Password)]
    [StringLength(100, MinimumLength = 8, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Password { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [DataType(DataType.Password)]
    [Compare(nameof(Password), ErrorMessage = "La confirmacion de contrasena no coincide.")]
    public string ConfirmPassword { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(80, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string RelacionClub { get; set; } = "Dueno";

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string NombreClub { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(80, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Pais { get; set; } = "Peru";

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string ProvinciaEstado { get; set; } = string.Empty;

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Ciudad { get; set; } = string.Empty;

    [Required(ErrorMessage = "Selecciona el departamento.")]
    [StringLength(2, ErrorMessage = "Codigo de departamento invalido.")]
    public string CodigoDepartamento { get; set; } = string.Empty;

    [Required(ErrorMessage = "Selecciona la provincia.")]
    [StringLength(4, ErrorMessage = "Codigo de provincia invalido.")]
    public string CodigoProvincia { get; set; } = string.Empty;

    [Required(ErrorMessage = "Selecciona el distrito.")]
    [StringLength(6, ErrorMessage = "Codigo de distrito invalido.")]
    public string CodigoUbigeo { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(250, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Direccion { get; set; } = string.Empty;

    public string? CaptchaTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(10, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string CaptchaCodigo { get; set; } = string.Empty;

}

public class AltasClubesIndexViewModel : ModuloBaseViewModel
{
    public int? Estado { get; set; }
    public List<AltaClubItemViewModel> Solicitudes { get; set; } = new();
}

public class AltaClubItemViewModel
{
    public int Id { get; set; }
    public string CodigoSolicitud { get; set; } = string.Empty;
    public string NombreContacto { get; set; } = string.Empty;
    public string Telefono { get; set; } = string.Empty;
    public string Correo { get; set; } = string.Empty;
    public string RelacionClub { get; set; } = string.Empty;
    public string NombreClub { get; set; } = string.Empty;
    public string Pais { get; set; } = string.Empty;
    public string ProvinciaEstado { get; set; } = string.Empty;
    public string Ciudad { get; set; } = string.Empty;
    public string Direccion { get; set; } = string.Empty;
    public int Estado { get; set; }
    public string? ComentarioGestion { get; set; }
    public int? NegocioId { get; set; }
    public int? SedeId { get; set; }
    public DateTime FechaRegistro { get; set; }
    public DateTime? FechaGestion { get; set; }
}

public class AltaClubGestionFormViewModel
{
    public int Id { get; set; }

    [StringLength(300, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? ComentarioGestion { get; set; }
}
