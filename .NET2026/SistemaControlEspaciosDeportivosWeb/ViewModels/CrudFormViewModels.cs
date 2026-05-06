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

    [Required(ErrorMessage = "Debes seleccionar un distrito valido.")]
    [StringLength(6, ErrorMessage = "El codigo ubigeo debe tener 6 caracteres.")]
    public string CodigoUbigeo { get; set; } = string.Empty;

    [StringLength(2)]
    public string? CodigoDepartamento { get; set; }

    [StringLength(4)]
    public string? CodigoProvincia { get; set; }

    [StringLength(2000, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? ConsideracionesReserva { get; set; }

    [Range(-90, 90, ErrorMessage = "La latitud debe estar entre -90 y 90.")]
    public decimal? Latitud { get; set; }

    [Range(-180, 180, ErrorMessage = "La longitud debe estar entre -180 y 180.")]
    public decimal? Longitud { get; set; }

    [StringLength(200, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? GooglePlaceId { get; set; }

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? GoogleDepartamento { get; set; }

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? GoogleProvincia { get; set; }

    [StringLength(120, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? GoogleDistrito { get; set; }

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

    [StringLength(500, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [Url(ErrorMessage = "Ingresa una URL valida.")]
    public string? FacebookUrl { get; set; }

    [StringLength(500, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [Url(ErrorMessage = "Ingresa una URL valida.")]
    public string? InstagramUrl { get; set; }

    [StringLength(500, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    [Url(ErrorMessage = "Ingresa una URL valida.")]
    public string? TwitterUrl { get; set; }

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
    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
    public List<SedeSerieDocumentoConfigItemViewModel> SeriesDocumentoConfig { get; set; } = new();
}

public class SedeSerieDocumentoConfigItemViewModel
{
    public string CodigoSunat { get; set; } = string.Empty;
    public string NombreDocumento { get; set; } = string.Empty;
    public bool Tributario { get; set; }
    public int? NegocioSerieId { get; set; }
    public List<int> NegocioSeriesIds { get; set; } = new();
    public string? SerieSeleccionada { get; set; }
    public bool PermiteMultiplesSeries =>
        string.Equals(CodigoSunat, "07", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(CodigoSunat, "08", StringComparison.OrdinalIgnoreCase);
    public List<SelectListItem> SeriesDisponibles { get; set; } = new();
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
    public bool AdministracionPrivada { get; set; }
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
    public bool PuedeModificarPrecio { get; set; }

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
    [StringLength(30, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? CodigoCupon { get; set; }

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

    [Range(1, int.MaxValue, ErrorMessage = "Selecciona una reserva valida.")]
    public int ReservaId { get; set; }

    [Range(0.01, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Monto { get; set; }

    public DateTime FechaPago { get; set; } = DateTime.Today;

    [Range(1, int.MaxValue, ErrorMessage = "Debes seleccionar una forma de pago.")]
    [Display(Name = "Forma de pago")]
    public int FormaPagoId { get; set; }

    [StringLength(50, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? NumeroOperacion { get; set; }

    [StringLength(300, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? Observacion { get; set; }

    public List<SelectListItem> Reservas { get; set; } = new();
    public List<SelectListItem> FormasPago { get; set; } = new();

    public string? ReservaTextoSeleccionada { get; set; }
    public string? Sede { get; set; }
    public string? Espacio { get; set; }
    public string? Cliente { get; set; }
    public DateOnly? FechaReserva { get; set; }
    public TimeOnly? HoraInicioReserva { get; set; }
    public TimeOnly? HoraFinReserva { get; set; }
    public decimal? TotalReserva { get; set; }
    public decimal? PagadoReserva { get; set; }
    public decimal? SaldoReserva { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public int PoliticaConfirmacionPago { get; set; }
    public decimal? PorcentajeAdelantoMinimo { get; set; }
    public List<PagoPrevioItemViewModel> PagosPrevios { get; set; } = new();
}

public class PagoPrevioItemViewModel
{
    public int PagoId { get; set; }
    public DateTime FechaPago { get; set; }
    public decimal Monto { get; set; }
    public string FormaPago { get; set; } = string.Empty;
    public string? NumeroOperacion { get; set; }
    public string? Observacion { get; set; }
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

    [Range(0, int.MaxValue, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
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
    public List<SelectListItem> TiposDocumentoComprobante { get; set; } = new();
    public string CodigoDocumentoComprobante { get; set; } = "03";
    public bool DocumentoTributario { get; set; } = true;
    public int? NegocioSerieId { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public bool EmisionComprobantesElectronicos { get; set; }
    public bool EmisionReciboInterno { get; set; }

    public string? ReservaCodigo { get; set; }
    public string? Sede { get; set; }
    public string? Espacio { get; set; }
    public string? Cliente { get; set; }
    public int? ClienteId { get; set; }
    public string? ClienteCorreo { get; set; }
    public string? ClienteTipoDocumento { get; set; }
    public string? ClienteNumeroDocumento { get; set; }
    public string? ClienteDireccionFiscal { get; set; }
    public string? ClienteCodigoDepartamento { get; set; }
    public string? ClienteCodigoProvincia { get; set; }
    public string? ClienteCodigoUbigeo { get; set; }
    public DateOnly? FechaReserva { get; set; }
    public TimeOnly? HoraInicioReserva { get; set; }
    public TimeOnly? HoraFinReserva { get; set; }
    public decimal? TotalReserva { get; set; }
    public decimal? PagadoReserva { get; set; }
    public decimal? SaldoReserva { get; set; }
    public int PorcentajeIgvConfigurado { get; set; } = 18;
    public List<PagoPrevioItemViewModel> PagosReserva { get; set; } = new();
    public List<SelectListItem> SeriesDocumento { get; set; } = new();
    public List<SelectListItem> TiposDocumentoIdentidad { get; set; } = new();
    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
    public bool EsNota { get; set; }
    public string? TipoNota { get; set; }
    public int? ComprobanteReferenciaId { get; set; }
    public string? ComprobanteReferenciaTipo { get; set; }
    public string? ComprobanteReferenciaSerie { get; set; }
    public int? ComprobanteReferenciaNumero { get; set; }
    public string? TipoNotaCodigoSunat { get; set; }
    public List<SelectListItem> TiposNotaSunat { get; set; } = new();
    public int CodigoDocumentoComprobantenb { get; set; }
    public int MonedaNubefact { get; set; } = 1;

    public bool EsEdicion { get; set; }
    public bool PuedeEditarDatosCliente => EsEdicion && Estado == EstadoComprobanteElectronico.PendienteEnvio;
}

public class ComprobanteReservaContextoViewModel
{
    public int ReservaId { get; set; }
    public string ReservaCodigo { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public string Cliente { get; set; } = string.Empty;
    public int? ClienteId { get; set; }
    public string? ClienteCorreo { get; set; }
    public string? ClienteTipoDocumento { get; set; }
    public string? ClienteNumeroDocumento { get; set; }
    public string? ClienteDireccionFiscal { get; set; }
    public string? ClienteCodigoUbigeo { get; set; }
    public string? ClienteCodigoDepartamento { get; set; }
    public string? ClienteCodigoProvincia { get; set; }
    public DateOnly FechaReserva { get; set; }
    public TimeOnly HoraInicioReserva { get; set; }
    public TimeOnly HoraFinReserva { get; set; }
    public decimal TotalReserva { get; set; }
    public decimal TotalPagado { get; set; }
    public decimal SaldoPendiente { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public int PorcentajeIgvConfigurado { get; set; } = 18;
    public List<PagoPrevioItemViewModel> PagosReserva { get; set; } = new();
    public List<SelectListItem> DocumentosDisponibles { get; set; } = new();
    public List<SelectListItem> SeriesDisponibles { get; set; } = new();
}

public class ComprobanteVisualizacionViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public int ReservaId { get; set; }
    public int TipoComprobante { get; set; }
    public string CodigoDocumentoComprobante { get; set; } = string.Empty;
    public string TipoDocumentoNombre { get; set; } = string.Empty;
    public bool EsTributario { get; set; }
    public string Serie { get; set; } = string.Empty;
    public int Numero { get; set; }
    public DateTime FechaEmision { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public decimal SubTotal { get; set; }
    public decimal Igv { get; set; }
    public decimal Total { get; set; }
    public int PorcentajeIgv { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string? NegocioRazonSocial { get; set; }
    public string? NegocioDireccionFiscal { get; set; }
    public string? NegocioDistrito { get; set; }
    public string? NegocioProvincia { get; set; }
    public string? NegocioDepartamento { get; set; }
    public string? NegocioDocumento { get; set; }
    public string ClienteNombre { get; set; } = string.Empty;
    public string? ClienteDocumento { get; set; }
    public string? ClienteDireccion { get; set; }
    public string? ClienteDistrito { get; set; }
    public string? ClienteProvincia { get; set; }
    public string? ClienteDepartamento { get; set; }
    public string? ClienteCorreo { get; set; }
    public string SedeNombre { get; set; } = string.Empty;
    public string EspacioNombre { get; set; } = string.Empty;
    public DateOnly FechaReserva { get; set; }
    public TimeOnly HoraInicioReserva { get; set; }
    public TimeOnly HoraFinReserva { get; set; }
    public string? UrlDescargaProveedor { get; set; }
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

public class CuponFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;
    public int? SedeId { get; set; }
    public int? EspacioDeportivoId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(30, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string CodigoCupon { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(150, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string Nombre { get; set; } = string.Empty;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [StringLength(20, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string TipoDescuento { get; set; } = "PORCENTAJE";

    [Range(0.01, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal ValorDescuento { get; set; }

    [Range(1, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int CantidadMaxUsos { get; set; } = 1;

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public DateOnly FechaInicio { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    public DateOnly FechaFin { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(30));

    public bool Activo { get; set; } = true;

    public List<SelectListItem> Sedes { get; set; } = new();
    public List<SelectListItem> Espacios { get; set; } = new();
}
