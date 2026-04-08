using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Models;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class SedeFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(150, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Nombre { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(250, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Direccion { get; set; } = string.Empty;

    [StringLength(2000, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? ConsideracionesReserva { get; set; }

    [Range(-90, 90, ErrorMessage = "La latitud debe estar entre -90 y 90.")]
    public decimal? Latitud { get; set; }

    [Range(-180, 180, ErrorMessage = "La longitud debe estar entre -180 y 180.")]
    public decimal? Longitud { get; set; }

    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? GooglePlaceId { get; set; }

    [StringLength(500, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [Url(ErrorMessage = "Ingresa una URL valida.")]
    public string? GoogleMapsUrl { get; set; }

    [StringLength(500, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [Url(ErrorMessage = "Ingresa una URL valida.")]
    public string? FotoPrincipalUrl { get; set; }

    public string? FotosUrlsCsv { get; set; }
    public List<string> FotosUrls { get; set; } = new();
    public List<string> FotosEliminarUrls { get; set; } = new();
    public List<IFormFile>? ImagenesArchivos { get; set; }

    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Telefono { get; set; }

    [RegularExpression(@"^\+\d{1,4}$", ErrorMessage = "Selecciona un codigo de pais valido.")]
    public string TelefonoCodigoPais { get; set; } = "+51";

    [RegularExpression(@"^$|^\d{6,15}$", ErrorMessage = "Ingresa un numero telefonico valido (solo digitos).")]
    public string? TelefonoNumeroLocal { get; set; }

    public bool Activo { get; set; } = true;

    public List<int> ServiciosSeleccionados { get; set; } = new();
    public List<SelectListItem> ServiciosDisponibles { get; set; } = new();

    public bool NotificacionesActivas { get; set; } = true;

    [Range(5, 1440, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int MinutosAnticipacionRecordatorio { get; set; } = 90;

    [Range(0, 240, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int MinutosToleranciaNoShow { get; set; } = 30;

    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [EmailAddress(ErrorMessage = "Ingresa un correo electronico valido.")]
    public string? CorreoNotificacion { get; set; }

    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? WhatsappContacto { get; set; }

    [RegularExpression(@"^\+\d{1,4}$", ErrorMessage = "Selecciona un codigo de pais valido.")]
    public string WhatsappCodigoPais { get; set; } = "+51";

    [RegularExpression(@"^$|^\d{6,15}$", ErrorMessage = "Ingresa un numero de WhatsApp valido (solo digitos).")]
    public string? WhatsappNumeroLocal { get; set; }

    public bool PermiteChatWhatsapp { get; set; }

    public bool AtiendeLunes { get; set; } = true;
    public bool AtiendeMartes { get; set; } = true;
    public bool AtiendeMiercoles { get; set; } = true;
    public bool AtiendeJueves { get; set; } = true;
    public bool AtiendeViernes { get; set; } = true;
    public bool AtiendeSabado { get; set; } = true;
    public bool AtiendeDomingo { get; set; } = true;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraApertura { get; set; } = new(8, 0);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraCierre { get; set; } = new(23, 0);

    public string? FechasInhabilitadasCsv { get; set; }
    public List<DateOnly> FechasInhabilitadas { get; set; } = new();
    public List<SelectListItem> CodigosPais { get; set; } = new();
}

public class EspacioFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Range(1, int.MaxValue, ErrorMessage = "Debes seleccionar una sede.")]
    public int SedeId { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Debes seleccionar un deporte.")]
    public int TipoDeporteId { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Debes seleccionar un tipo de suelo.")]
    public int TipoSueloId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Codigo { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(150, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Nombre { get; set; } = string.Empty;

    [Range(1, 200, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int Capacidad { get; set; } = 10;

    public bool TieneIluminacion { get; set; }
    public bool Techada { get; set; }
    public EstadoEspacioDeportivo Estado { get; set; } = EstadoEspacioDeportivo.Activo;
    public string? TarifasJson { get; set; }
    public List<EspacioTarifaRangoViewModel> Tarifas { get; set; } = new();
    public int? MonedaIdConfigurada { get; set; }
    public string MonedaEtiqueta { get; set; } = string.Empty;
    public bool PuedeEditarTarifas { get; set; }

    public List<SelectListItem> Sedes { get; set; } = new();
    public List<SelectListItem> TiposDeporte { get; set; } = new();
    public List<SelectListItem> TiposSuelo { get; set; } = new();
    public List<SelectListItem> TarifaDiasSemana { get; set; } = new();
}

public class EspacioTarifaRangoViewModel
{
    public int DiaSemana { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public decimal Precio { get; set; }
}

public class ReservaFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public int EspacioDeportivoId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public int ClienteId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public DateOnly Fecha { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraInicio { get; set; } = new(18, 0);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraFin { get; set; } = new(19, 0);

    [Range(0, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Total { get; set; }

    [Range(0, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Adelanto { get; set; }

    public EstadoReserva Estado { get; set; } = EstadoReserva.Pendiente;
    public bool RegistrarPago { get; set; }
    public int? FormaPagoId { get; set; }
    public DateTime? FechaPago { get; set; }

    [StringLength(50, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? NumeroOperacion { get; set; }

    [StringLength(500, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Comentario { get; set; }

    public List<SelectListItem> Espacios { get; set; } = new();
    public List<SelectListItem> Clientes { get; set; } = new();
    public List<SelectListItem> FormasPago { get; set; } = new();
}

public class BloqueoHorarioFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public int EspacioDeportivoId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public DateOnly Fecha { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraInicio { get; set; } = new(8, 0);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraFin { get; set; } = new(9, 0);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(250, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Motivo { get; set; } = string.Empty;

    public List<SelectListItem> Espacios { get; set; } = new();
}

public class PagoFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public int ReservaId { get; set; }

    [Range(0.01, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Monto { get; set; }

    public DateTime FechaPago { get; set; } = DateTime.Now;

    [Range(1, int.MaxValue, ErrorMessage = "Debes seleccionar una forma de pago.")]
    [Display(Name = "Forma de pago")]
    public int FormaPagoId { get; set; }

    [StringLength(50, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? NumeroOperacion { get; set; }

    [StringLength(300, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Observacion { get; set; }

    public List<SelectListItem> Reservas { get; set; } = new();
    public List<SelectListItem> FormasPago { get; set; } = new();
}

public class ComprobanteFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public int ReservaId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TipoComprobante TipoComprobante { get; set; } = TipoComprobante.Boleta;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(4, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Serie { get; set; } = "B001";

    [Range(1, int.MaxValue, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int Numero { get; set; }

    public DateTime FechaEmision { get; set; } = DateTime.Now;
    public TipoMoneda TipoMoneda { get; set; } = TipoMoneda.PEN;

    [Range(0, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal SubTotal { get; set; }

    [Range(0, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Igv { get; set; }

    [Range(0, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Total { get; set; }

    public EstadoComprobanteElectronico Estado { get; set; } = EstadoComprobanteElectronico.PendienteEnvio;

    public List<SelectListItem> Reservas { get; set; } = new();
}

public class ClienteFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string NombresORazonSocial { get; set; } = string.Empty;

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Nombres { get; set; }

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Apellidos { get; set; }

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? NombreEquipo { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string TipoDocumento { get; set; } = "0";

    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string NumeroDocumento { get; set; } = string.Empty;

    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Telefono { get; set; }

    [RegularExpression(@"^\+\d{1,4}$", ErrorMessage = "Selecciona un codigo de pais valido.")]
    public string TelefonoCodigoPais { get; set; } = "+51";

    [RegularExpression(@"^$|^\d{6,15}$", ErrorMessage = "Ingresa un numero telefonico valido (solo digitos).")]
    public string? TelefonoNumeroLocal { get; set; }

    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [EmailAddress(ErrorMessage = "Ingresa un correo electronico valido.")]
    public string? Correo { get; set; }

    [StringLength(250, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? DireccionFiscal { get; set; }
    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    [StringLength(6, ErrorMessage = "El campo {0} debe tener 6 caracteres.")]
    public string? CodigoUbigeo { get; set; }

    public bool Activo { get; set; } = true;
    public List<SelectListItem> CodigosPais { get; set; } = new();
    public List<SelectListItem> TiposDocumento { get; set; } = new();
    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
}

public class PromocionFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    public int? SedeId { get; set; }
    public int? EspacioDeportivoId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(150, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Nombre { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public DateOnly FechaInicio { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public DateOnly FechaFin { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraInicio { get; set; } = new(8, 0);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public TimeOnly HoraFin { get; set; } = new(10, 0);

    [Range(0, 100, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal PorcentajeDescuento { get; set; }

    public bool Activo { get; set; } = true;

    public List<SelectListItem> Sedes { get; set; } = new();
    public List<SelectListItem> Espacios { get; set; } = new();
}
