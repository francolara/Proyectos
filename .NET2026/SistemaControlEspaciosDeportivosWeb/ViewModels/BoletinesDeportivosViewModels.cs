using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class BoletinDeportivoPublicoItemViewModel
{
    public int IdBoletin { get; set; }
    public string? Titulo { get; set; }
    public string? Descripcion { get; set; }
    public string ImagenUrl { get; set; } = string.Empty;
    public DateOnly FechaEvento { get; set; }
    public string CodigoUbigeo { get; set; } = string.Empty;
    public string Departamento { get; set; } = string.Empty;
    public string Provincia { get; set; } = string.Empty;
    public string Distrito { get; set; } = string.Empty;
    public string? Zona { get; set; }
    public string TipoRegistro { get; set; } = "U";
    public DateTime FechaCreacion { get; set; }
}

public class BoletinDeportivoUsuarioItemViewModel : BoletinDeportivoPublicoItemViewModel
{
    public bool Activo { get; set; }
}

public class BoletinDeportivoDetalleViewModel : BoletinDeportivoUsuarioItemViewModel
{
    public string UsuarioId { get; set; } = string.Empty;
    public int? PerfilPublicoId { get; set; }
    public string CodigoDepartamento { get; set; } = string.Empty;
    public string CodigoProvincia { get; set; } = string.Empty;
    public string UsuarioCreacion { get; set; } = string.Empty;
    public DateTime? FechaActualizacion { get; set; }
    public string? UsuarioActualizacion { get; set; }
}

public class BoletinDeportivoAdminItemViewModel : BoletinDeportivoUsuarioItemViewModel
{
    public string UsuarioId { get; set; } = string.Empty;
    public string? CorreoAutor { get; set; }
    public string NombreAutor { get; set; } = string.Empty;
}

public class BoletinDeportivoGuardarViewModel
{
    public int? IdBoletin { get; set; }
    public string UsuarioId { get; set; } = string.Empty;

    [StringLength(160, ErrorMessage = "El titulo no puede exceder 160 caracteres.")]
    public string? Titulo { get; set; }

    [StringLength(500, ErrorMessage = "La descripcion no puede exceder 500 caracteres.")]
    public string? Descripcion { get; set; }

    [Required(ErrorMessage = "La fecha del evento es obligatoria.")]
    public DateOnly? FechaEvento { get; set; }

    [Required(ErrorMessage = "Debes seleccionar el distrito del evento.")]
    [StringLength(6, ErrorMessage = "El codigo ubigeo debe tener 6 caracteres.")]
    public string CodigoUbigeo { get; set; } = string.Empty;

    [StringLength(500, ErrorMessage = "La URL de la imagen no puede exceder 500 caracteres.")]
    public string ImagenUrl { get; set; } = string.Empty;

    public IFormFile? ImagenArchivo { get; set; }
    public bool Activo { get; set; } = true;
    public bool EsAdministradorCarga { get; set; }
    public string TipoRegistro { get; set; } = "U";
    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    public string? Zona { get; set; }
    public List<SelectListItem> Departamentos { get; set; } = new();
    public List<SelectListItem> Provincias { get; set; } = new();
    public List<SelectListItem> Distritos { get; set; } = new();
    public List<SelectListItem> Zonas { get; set; } = new();
}

public class BoletinesPublicosIndexViewModel
{
    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    public string? CodigoUbigeo { get; set; }
    public string? Zona { get; set; }
    public int Anio { get; set; }
    public int Mes { get; set; }
    public List<SelectListItem> Departamentos { get; set; } = new();
    public List<SelectListItem> Provincias { get; set; } = new();
    public List<SelectListItem> Distritos { get; set; } = new();
    public List<SelectListItem> Zonas { get; set; } = new();
    public List<SelectListItem> Anios { get; set; } = new();
    public List<SelectListItem> Meses { get; set; } = new();
    public List<BoletinDeportivoPublicoItemViewModel> Boletines { get; set; } = new();
}
