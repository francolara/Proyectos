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

    [RegularExpression(@"^\+\d{1,4}$", ErrorMessage = "Selecciona un codigo de pais valido.")]
    public string TelefonoCodigoPais { get; set; } = "+51";

    [RegularExpression(@"^$|^\d{6,15}$", ErrorMessage = "Ingresa un numero telefonico valido (solo digitos).")]
    public string? TelefonoNumeroLocal { get; set; }

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

    public bool BuscarDesafios { get; set; }

    public int? IdDeporteDesafio { get; set; }

    public int? IdNivelDesafio { get; set; }

    [StringLength(500)]
    public string? ObservacionDesafio { get; set; }

    [StringLength(1000)]
    public string? DetalleEquipo { get; set; }

    [StringLength(6)]
    public string? CodigoUbigeoEquipo { get; set; }

    [StringLength(2)]
    public string? CodigoDepartamentoEquipo { get; set; }

    [StringLength(4)]
    public string? CodigoProvinciaEquipo { get; set; }

    [StringLength(30)]
    public string? WhatsappEquipo { get; set; }

    [RegularExpression(@"^\+\d{1,4}$", ErrorMessage = "Selecciona un codigo de pais valido.")]
    public string WhatsappCodigoPais { get; set; } = "+51";

    [RegularExpression(@"^$|^\d{6,15}$", ErrorMessage = "Ingresa un numero de WhatsApp valido (solo digitos).")]
    public string? WhatsappNumeroLocal { get; set; }

    public List<SelectListItem> TiposDocumento { get; set; } = new();
    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
    public List<SelectListItem> DepartamentosUbigeoEquipo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeoEquipo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeoEquipo { get; set; } = new();
    public List<SelectListItem> DeportesDesafio { get; set; } = new();
    public List<SelectListItem> NivelesDesafio { get; set; } = new();
    public List<SelectListItem> CodigosPais { get; set; } = new();
}

public class UsuarioPublicoReservaItemViewModel
{
    public int ReservaId { get; set; }
    public string CodigoReserva => $"R-{ReservaId:D6}";
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
    public string? SedeFacebookUrl { get; set; }
    public string? SedeInstagramUrl { get; set; }
    public string? SedeTwitterUrl { get; set; }
    public string? SedeMapaUrl { get; set; }
    public bool PuedeRegistrarResena { get; set; }
    public bool PuedeAgregarCalendario => EstadoTexto is not "Cancelada" and not "No Show";
    public string AliasResenaSugerido { get; set; } = string.Empty;
    public UsuarioPublicoResenaItemViewModel? Resena { get; set; }
}

public class UsuarioPublicoReservaCalendarioViewModel
{
    public int ReservaId { get; set; }
    public string CodigoReserva => $"R-{ReservaId:D6}";
    public int EstadoId { get; set; }
    public string EstadoTexto { get; set; } = string.Empty;
    public string NegocioNombre { get; set; } = string.Empty;
    public string SedeNombre { get; set; } = string.Empty;
    public string EspacioNombre { get; set; } = string.Empty;
    public string? SedeDireccion { get; set; }
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
}

public class UsuarioPublicoResenaItemViewModel
{
    public int ResenaId { get; set; }
    public int ReservaId { get; set; }
    public int EspacioDeportivoId { get; set; }
    public string AliasPublico { get; set; } = string.Empty;
    public string Comentario { get; set; } = string.Empty;
    public bool Activo { get; set; } = true;
    public string? Respuesta { get; set; }
    public DateTime FechaCreacion { get; set; }
}

public class UsuarioPublicoResenaGuardarViewModel
{
    [Required(ErrorMessage = "La reserva es obligatoria.")]
    public int ReservaId { get; set; }

    public int Pagina { get; set; } = 1;

    [Required(ErrorMessage = "El alias visible es obligatorio.")]
    [StringLength(120, ErrorMessage = "El alias visible no puede exceder 120 caracteres.")]
    public string AliasPublico { get; set; } = string.Empty;

    [Required(ErrorMessage = "El comentario de la resena es obligatorio.")]
    [StringLength(800, ErrorMessage = "El comentario no puede exceder 800 caracteres.")]
    public string Comentario { get; set; } = string.Empty;
}

public class PerfilPublicoIndexViewModel
{
    public UsuarioPublicoPerfilViewModel Perfil { get; set; } = new();
    public List<UsuarioPublicoReservaItemViewModel> Reservas { get; set; } = new();
    public BoletinDeportivoGuardarViewModel BoletinForm { get; set; } = new();
    public List<BoletinDeportivoUsuarioItemViewModel> Boletines { get; set; } = new();
    public int PaginaReservas { get; set; } = 1;
    public int TamanoPaginaReservas { get; set; } = 6;
    public int TotalReservas { get; set; }
    public int TotalPaginasReservas { get; set; } = 1;
    public int PaginaBoletines { get; set; } = 1;
    public int TamanoPaginaBoletines { get; set; } = 5;
    public int TotalBoletines { get; set; }
    public int TotalPaginasBoletines { get; set; } = 1;
}
