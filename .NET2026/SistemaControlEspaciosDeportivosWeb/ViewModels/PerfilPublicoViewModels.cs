using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class UsuarioPublicoPerfilViewModel
{
    public int Id { get; set; }
    public string UsuarioId { get; set; } = string.Empty;

    [Required(ErrorMessage = "El tipo de documento es obligatorio.")]
    [StringLength(20)]
    public string TipoDocumento { get; set; } = "0";

    [StringLength(20)]
    public string? NumeroDocumento { get; set; }

    [Required(ErrorMessage = "Los nombres son obligatorios.")]
    [StringLength(120)]
    public string Nombres { get; set; } = string.Empty;

    [Required(ErrorMessage = "Los apellidos son obligatorios.")]
    [StringLength(120)]
    public string Apellidos { get; set; } = string.Empty;

    [StringLength(120)]
    public string? NombreEquipo { get; set; }

    [StringLength(30)]
    public string? Telefono { get; set; }

    [StringLength(200)]
    [EmailAddress(ErrorMessage = "Ingresa un correo electronico valido.")]
    public string? Correo { get; set; }

    public DateOnly? FechaNacimiento { get; set; }

    [StringLength(6)]
    public string? CodigoUbigeo { get; set; }

    [StringLength(2)]
    public string? CodigoDepartamento { get; set; }

    [StringLength(4)]
    public string? CodigoProvincia { get; set; }

    public List<SelectListItem> TiposDocumento { get; set; } = new();
    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
}

public class UsuarioPublicoReservaItemViewModel
{
    public int ReservaId { get; set; }
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string EstadoTexto { get; set; } = string.Empty;
    public decimal Total { get; set; }
    public decimal Adelanto { get; set; }
    public decimal SaldoPendiente { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string SedeNombre { get; set; } = string.Empty;
    public string EspacioNombre { get; set; } = string.Empty;
    public string? SedeDireccion { get; set; }
    public string? SedeTelefono { get; set; }
    public string? SedeWhatsapp { get; set; }
}

public class PerfilPublicoIndexViewModel
{
    public UsuarioPublicoPerfilViewModel Perfil { get; set; } = new();
    public List<UsuarioPublicoReservaItemViewModel> Reservas { get; set; } = new();
}
