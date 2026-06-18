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
    [StringLength(20)]
    public string NumeroDocumento { get; set; } = string.Empty;

    [StringLength(100)]
    public string? ApellidoPaterno { get; set; }

    [StringLength(100)]
    public string? ApellidoMaterno { get; set; }

    [StringLength(150)]
    public string? Nombres { get; set; }

    [StringLength(200)]
    public string? RazonSocial { get; set; }

    [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
    [StringLength(200)]
    public string? CorreoElectronico { get; set; }

    [StringLength(50)]
    public string? Telefono { get; set; }

    [StringLength(250)]
    public string? Direccion { get; set; }

    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    public string? CodigoUbigeo { get; set; }
    public bool EsCliente { get; set; }
    public bool EsProveedor { get; set; }
    public bool Estado { get; set; } = true;
}
