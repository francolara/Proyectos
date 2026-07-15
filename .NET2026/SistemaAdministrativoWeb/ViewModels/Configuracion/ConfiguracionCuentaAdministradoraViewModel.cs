using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Configuracion;

public sealed class ConfiguracionCuentaAdministradoraViewModel
{
    public int IdCuentaAdministradora { get; set; }

    public string CodigoCuenta { get; set; } = string.Empty;

    public string NombreCuenta { get; set; } = string.Empty;

    public string? CorreoPrincipal { get; set; }

    public string? TelefonoPrincipal { get; set; }

    [Display(Name = "Nombre Completo")]
    [StringLength(180, ErrorMessage = "El nombre del responsable principal no puede exceder 180 caracteres.")]
    public string? NombreResponsablePrincipal { get; set; }

    [Display(Name = "Correo")]
    [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
    [StringLength(256, ErrorMessage = "El correo administrativo no puede exceder 256 caracteres.")]
    public string? CorreoAdministrativo { get; set; }

    [Display(Name = "Telefono")]
    [StringLength(30, ErrorMessage = "El telefono administrativo no puede exceder 30 caracteres.")]
    public string? TelefonoAdministrativo { get; set; }

    [Display(Name = "Empresa predeterminada")]
    public int? IdEmpresaPredeterminada { get; set; }

    [Display(Name = "Observacion administrativa")]
    [StringLength(400, ErrorMessage = "La observacion administrativa no puede exceder 400 caracteres.")]
    public string? ObservacionAdministrativa { get; set; }

    [Required(ErrorMessage = "Seleccione el comprobante preferido.")]
    [Display(Name = "Tipo Comprobante")]
    public string TipoComprobantePreferido { get; set; } = "BOLETA";

    [Required(ErrorMessage = "Seleccione el tipo de documento.")]
    [Display(Name = "Tipo de documento")]
    public string TipoDocumentoFacturacion { get; set; } = "DNI";

    [Display(Name = "Numero de documento")]
    [StringLength(20, ErrorMessage = "El numero de documento no puede exceder 20 caracteres.")]
    public string? NumeroDocumento { get; set; }

    [Display(Name = "Nombre completo")]
    [StringLength(200, ErrorMessage = "El nombre completo no puede exceder 200 caracteres.")]
    public string? NombreFacturacion { get; set; }

    [Display(Name = "Razon social")]
    [StringLength(200, ErrorMessage = "La razon social no puede exceder 200 caracteres.")]
    public string? RazonSocialFacturacion { get; set; }

    [Display(Name = "Correo de facturacion")]
    [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
    [StringLength(256, ErrorMessage = "El correo de facturacion no puede exceder 256 caracteres.")]
    public string? CorreoFacturacion { get; set; }

    [Display(Name = "Telefono de facturacion")]
    [StringLength(30, ErrorMessage = "El telefono de facturacion no puede exceder 30 caracteres.")]
    public string? TelefonoFacturacion { get; set; }

    [Display(Name = "Direccion fiscal")]
    [StringLength(250, ErrorMessage = "La direccion fiscal no puede exceder 250 caracteres.")]
    public string? DireccionFiscal { get; set; }

    [Display(Name = "Ubigeo")]
    [StringLength(6, ErrorMessage = "El ubigeo no puede exceder 6 caracteres.")]
    public string? Ubigeo { get; set; }

    [Display(Name = "Distrito")]
    [StringLength(100, ErrorMessage = "El distrito no puede exceder 100 caracteres.")]
    public string? Distrito { get; set; }

    [Display(Name = "Provincia")]
    [StringLength(100, ErrorMessage = "La provincia no puede exceder 100 caracteres.")]
    public string? Provincia { get; set; }

    [Display(Name = "Departamento")]
    [StringLength(100, ErrorMessage = "El departamento no puede exceder 100 caracteres.")]
    public string? Departamento { get; set; }

    [Display(Name = "Observacion de facturacion")]
    [StringLength(400, ErrorMessage = "La observacion de facturacion no puede exceder 400 caracteres.")]
    public string? ObservacionFacturacion { get; set; }

    public IReadOnlyCollection<ConfiguracionEmpresaItemViewModel> EmpresasDisponibles { get; set; } = [];
}

public sealed class ConfiguracionEmpresaItemViewModel
{
    public int IdEmpresa { get; set; }
    public string CodigoEmpresa { get; set; } = string.Empty;
    public string RazonSocial { get; set; } = string.Empty;
}

public sealed class FacturacionPadronLookupResultViewModel
{
    public bool Encontrado { get; set; }
    public string? NumeroDocumento { get; set; }
    public string? NombreFacturacion { get; set; }
    public string? RazonSocialFacturacion { get; set; }
    public string? DireccionFiscal { get; set; }
    public string? Ubigeo { get; set; }
    public string? Distrito { get; set; }
    public string? Provincia { get; set; }
    public string? Departamento { get; set; }
    public string? Mensaje { get; set; }
}
