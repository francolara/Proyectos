using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PersonaFormViewModel
{
    public int? IdPersona { get; set; }

    [Required(ErrorMessage = "Seleccione el tipo de persona.")]
    public string TipoPersona { get; set; } = "N";

    [Required(ErrorMessage = "Seleccione el tipo de documento.")]
    public string TipoDocumento { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el numero de documento.")]
    [StringLength(20, ErrorMessage = "El numero de documento no puede exceder 20 caracteres.")]
    public string NumeroDocumento { get; set; } = string.Empty;

    [StringLength(100, ErrorMessage = "El apellido paterno no puede exceder 100 caracteres.")]
    public string? ApellidoPaterno { get; set; }

    [StringLength(100, ErrorMessage = "El apellido materno no puede exceder 100 caracteres.")]
    public string? ApellidoMaterno { get; set; }

    [StringLength(150, ErrorMessage = "Los nombres no pueden exceder 150 caracteres.")]
    public string? Nombres { get; set; }

    [StringLength(200, ErrorMessage = "La razon social no puede exceder 200 caracteres.")]
    public string? RazonSocial { get; set; }

    [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
    [StringLength(200, ErrorMessage = "El correo electronico no puede exceder 200 caracteres.")]
    public string? CorreoElectronico { get; set; }

    [StringLength(50, ErrorMessage = "El telefono no puede exceder 50 caracteres.")]
    public string? Telefono { get; set; }

    [StringLength(250, ErrorMessage = "La direccion no puede exceder 250 caracteres.")]
    public string? Direccion { get; set; }

    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    public string? CodigoUbigeo { get; set; }
    public bool EsCliente { get; set; }
    public bool EsProveedor { get; set; }
    public bool Estado { get; set; } = true;
}
