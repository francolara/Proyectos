using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class HomeIndexViewModel
{
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    public string? CodigoUbigeo { get; set; }
    public int? TipoDeporteId { get; set; }
    public int? NegocioId { get; set; }
    public bool? OmitirFechaHorario { get; set; }
    public bool BuscarCercaDeMi { get; set; }
    public decimal? LatitudUsuario { get; set; }
    public decimal? LongitudUsuario { get; set; }
    public decimal? RadioKm { get; set; }
    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
    public List<SelectListItem> Negocios { get; set; } = new();
    public List<WebBannerPublicoViewModel> Banners { get; set; } = new();
    public List<PopupPromocionPublicoViewModel> PopupPromociones { get; set; } = new();
    public List<SedePublicaViewModel> Sedes { get; set; } = new();
    public List<TipoDeportePublicoViewModel> TiposDeporte { get; set; } = new();
    public List<EspacioDisponibleViewModel> Disponibles { get; set; } = new();
    public string? MensajeSolicitud { get; set; }
    public int PaginaActual { get; set; } = 1;
    public int TamanoPagina { get; set; } = 9;
    public int TotalResultados { get; set; }
    public int TotalPaginas { get; set; }
    public PlataformaPortalConfigViewModel PortalConfig { get; set; } = new();
    public PopupPromocionConfigViewModel PopupPromocionesConfig { get; set; } = new();
}

public class SedePublicaViewModel
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public int? NegocioId { get; set; }
    public string? NegocioNombre { get; set; }
    public string? Servicios { get; set; }
    public string Direccion { get; set; } = string.Empty;
    public string? ConsideracionesReserva { get; set; }
    public string? Telefono { get; set; }
    public string? WhatsappContacto { get; set; }
    public bool PermiteChatWhatsapp { get; set; }
    public string? FacebookUrl { get; set; }
    public string? InstagramUrl { get; set; }
    public string? TwitterUrl { get; set; }
    public decimal? Latitud { get; set; }
    public decimal? Longitud { get; set; }
    public string? GoogleMapsUrl { get; set; }
    public string? FotoPrincipalUrl { get; set; }
    public List<string> FotosAlternativas { get; set; } = new();
    public string? CodigoUbigeoNegocio { get; set; }
    public string? CodigoDepartamentoNegocio { get; set; }
    public string? CodigoProvinciaNegocio { get; set; }
}

public class TipoDeportePublicoViewModel
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
}

public class WebBannerPublicoViewModel
{
    public int Id { get; set; }
    public string Titulo { get; set; } = string.Empty;
    public string? Subtitulo { get; set; }
    public string? Descripcion { get; set; }
    public string? BotonTexto { get; set; }
    public string? BotonUrl { get; set; }
    public string ImagenUrl { get; set; } = string.Empty;
    public string? ImagenUrlMobile { get; set; }
    public int Orden { get; set; }
}

public class EspacioDisponibleViewModel
{
    public int EspacioDeportivoId { get; set; }
    public string NombreEspacio { get; set; } = string.Empty;
    public string Codigo { get; set; } = string.Empty;
    public int? SedeId { get; set; }
    public string SedeNombre { get; set; } = string.Empty;
    public string? SedeDireccion { get; set; }
    public string? SedeConsideracionesReserva { get; set; }
    public string? Departamento { get; set; }
    public string? Provincia { get; set; }
    public string? Distrito { get; set; }
    public string TipoDeporteNombre { get; set; } = string.Empty;
    public string? TipoSueloNombre { get; set; }
    public decimal? TarifaDesde { get; set; }
    public bool TieneIluminacion { get; set; }
    public bool Techada { get; set; }
    public string? CorreoNotificacion { get; set; }
    public string? TelefonoContacto { get; set; }
    public string? WhatsappContacto { get; set; }
    public bool PermiteChatWhatsapp { get; set; }
    public string? SedeMapaUrl { get; set; }
    public string? SedeFotoPrincipalUrl { get; set; }
    public List<string> SedeFotos { get; set; } = new();
    public decimal? DistanciaKm { get; set; }
    public string? NegocioNombreDestacado { get; set; }
    public string? TelefonoContactoResuelto { get; set; }
    public string? SedeMapaUrlResuelto { get; set; }
    public string? EnlaceWhatsappEspacio { get; set; }
    public List<string> SedeFotosConFallback { get; set; } = new();
    public int? NegocioIdCotizacion { get; set; }
    public string? SedeFacebookUrl { get; set; }
    public string? SedeInstagramUrl { get; set; }
    public string? SedeTwitterUrl { get; set; }
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
    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Nombres { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Apellidos { get; set; } = string.Empty;

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? NombreEquipo { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string TipoDocumento { get; set; } = "0";

    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? NumeroDocumento { get; set; }

    [StringLength(30, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Telefono { get; set; }

    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [EmailAddress(ErrorMessage = "Ingresa un correo electronico valido.")]
    public string? Correo { get; set; }

    [StringLength(300, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Comentario { get; set; }
    [StringLength(30, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? CodigoCupon { get; set; }

    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    public string? CodigoUbigeo { get; set; }
    public int? TipoDeporteId { get; set; }
    public int? NegocioId { get; set; }
    public bool? OmitirFechaHorario { get; set; }
    public string? UsuarioId { get; set; }
    [ValidateNever]
    public List<SelectListItem> TiposDocumentoIdentidad { get; set; } = new();
}

public class ReservaPublicaPageViewModel
{
    public int NegocioId { get; set; }
    public EspacioDisponibleViewModel Espacio { get; set; } = new();
    public SedePublicaViewModel? Sede { get; set; }
    public SolicitudReservaPublicaFormViewModel Formulario { get; set; } = new();
    public ReservaCotizacionViewModel? Cotizacion { get; set; }
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
