using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class HomeIndexViewModel
{
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public int? SedeId { get; set; }
    public int? TipoDeporteId { get; set; }
    public List<SedePublicaViewModel> Sedes { get; set; } = new();
    public List<TipoDeportePublicoViewModel> TiposDeporte { get; set; } = new();
    public List<EspacioDisponibleViewModel> Disponibles { get; set; } = new();
    public string? MensajeSolicitud { get; set; }
}

public class SedePublicaViewModel
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string Direccion { get; set; } = string.Empty;
    public string? Telefono { get; set; }
    public string? WhatsappContacto { get; set; }
    public bool PermiteChatWhatsapp { get; set; }
    public decimal? Latitud { get; set; }
    public decimal? Longitud { get; set; }
    public string? GoogleMapsUrl { get; set; }
    public string? FotoPrincipalUrl { get; set; }
    public List<string> FotosAlternativas { get; set; } = new();
}

public class TipoDeportePublicoViewModel
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
}

public class EspacioDisponibleViewModel
{
    public int EspacioDeportivoId { get; set; }
    public string NombreEspacio { get; set; } = string.Empty;
    public string Codigo { get; set; } = string.Empty;
    public string SedeNombre { get; set; } = string.Empty;
    public string TipoDeporteNombre { get; set; } = string.Empty;
    public bool TieneIluminacion { get; set; }
    public bool Techada { get; set; }
    public string? WhatsappContacto { get; set; }
    public bool PermiteChatWhatsapp { get; set; }
}

public class SolicitudReservaPublicaFormViewModel
{
    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public int EspacioDeportivoId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public DateOnly Fecha { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraInicio { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraFin { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string NombreSolicitante { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(30, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Telefono { get; set; } = string.Empty;

    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [EmailAddress(ErrorMessage = "Ingresa un correo electronico valido.")]
    public string? Correo { get; set; }

    [StringLength(300, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Comentario { get; set; }

    public int? SedeId { get; set; }
    public int? TipoDeporteId { get; set; }
}

public class SolicitudPublicaDetalleViewModel
{
    public string CodigoSolicitud { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string NombreSolicitante { get; set; } = string.Empty;
    public string Telefono { get; set; } = string.Empty;
    public string? Correo { get; set; }
    public int Estado { get; set; }
    public string EstadoTexto { get; set; } = string.Empty;
    public int? ReservaId { get; set; }
    public DateTime FechaRegistro { get; set; }
}

public class SolicitudPublicaSeguimientoViewModel
{
    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string CodigoSolicitud { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(30, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Telefono { get; set; } = string.Empty;

    public SolicitudPublicaDetalleViewModel? Resultado { get; set; }
    public string? Mensaje { get; set; }
}

public class SolicitudNotificacionEmailViewModel
{
    public string CodigoSolicitud { get; set; } = string.Empty;
    public string NombreSolicitante { get; set; } = string.Empty;
    public string Correo { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public bool NotificadoCliente { get; set; }
}
