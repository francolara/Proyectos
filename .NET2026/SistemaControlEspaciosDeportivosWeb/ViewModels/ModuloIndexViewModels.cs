using Microsoft.AspNetCore.Mvc.Rendering;
using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class ModuloBaseViewModel
{
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;
    public int? SedeIdAsignada { get; set; }
    public bool EsAdministrador { get; set; }
    public string ModuloCodigo { get; set; } = string.Empty;
    public string ModuloNombre { get; set; } = string.Empty;
    public bool PuedeCrear { get; set; }
    public bool PuedeEditar { get; set; }
    public bool PuedeEliminar { get; set; }
    public string? Mensaje { get; set; }
}

public class ConfiguracionClubViewModel : ModuloBaseViewModel
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El nombre comercial es obligatorio.")]
    [StringLength(200, ErrorMessage = "El nombre comercial no puede superar los 200 caracteres.")]
    public string NombreComercial { get; set; } = string.Empty;

    [StringLength(200, ErrorMessage = "La razón social no puede superar los 200 caracteres.")]
    public string? RazonSocial { get; set; }

    [StringLength(20, ErrorMessage = "El numero de documento no puede superar los 20 caracteres.")]
    public string? NumeroDocumento { get; set; }

    [Required(ErrorMessage = "El tipo de documento es obligatorio.")]
    [StringLength(20, ErrorMessage = "El tipo de documento no puede superar los 20 caracteres.")]
    public string TipoDocumento { get; set; } = "1";

    [StringLength(250, ErrorMessage = "La direccion fiscal no puede superar los 250 caracteres.")]
    public string? DireccionFiscal { get; set; }
    public string? CodigoDepartamento { get; set; }
    public string? CodigoProvincia { get; set; }
    [StringLength(6, ErrorMessage = "El codigo ubigeo debe tener 6 caracteres.")]
    public string? CodigoUbigeo { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Debes seleccionar una moneda válida.")]
    public int MonedaId { get; set; } = 1;

    [Range(0, 2, ErrorMessage = "La politica de confirmacion no es valida.")]
    public int PoliticaConfirmacionPago { get; set; } = 0;

    [Range(typeof(decimal), "1", "100", ErrorMessage = "El porcentaje minimo debe ser un numero entero entre 1 y 100.")]
    public decimal? PorcentajeAdelantoMinimo { get; set; }
    public bool EmisionComprobantesElectronicos { get; set; }
    public bool EmisionReciboInterno { get; set; }

    [Range(0, 100, ErrorMessage = "El porcentaje de IGV debe estar entre 0 y 100.")]
    public int PorcentajeIgv { get; set; } = 18;

    public List<SelectListItem> TiposDocumento { get; set; } = new();
    public List<SelectListItem> Monedas { get; set; } = new();
    public List<SelectListItem> PoliticasConfirmacionPago { get; set; } = new();
    public List<SelectListItem> TiposDocumentoComprobanteTributarios { get; set; } = new();
    public List<SelectListItem> TiposDocumentoComprobanteNoTributarios { get; set; } = new();
    public List<SerieDocumentoComprobanteItemViewModel> SeriesDocumentoComprobante { get; set; } = new();
    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
}

public class SerieDocumentoComprobanteItemViewModel
{
    public int Id { get; set; }
    public string CodigoSunat { get; set; } = string.Empty;
    public string NombreDocumento { get; set; } = string.Empty;
    public bool Tributario { get; set; }
    public string Serie { get; set; } = string.Empty;
    public bool Activo { get; set; }
}

public class UbigeoLookupViewModel
{
    public string CodigoUbigeo { get; set; } = string.Empty;
    public string CodigoDepartamento { get; set; } = string.Empty;
    public string CodigoProvincia { get; set; } = string.Empty;
    public string Departamento { get; set; } = string.Empty;
    public string Provincia { get; set; } = string.Empty;
    public string Distrito { get; set; } = string.Empty;
}

public class SedesIndexViewModel : ModuloBaseViewModel
{
    public List<SedeItemViewModel> Sedes { get; set; } = new();
}

public class SedeItemViewModel
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string Direccion { get; set; } = string.Empty;
    public string Servicios { get; set; } = string.Empty;
    public bool NotificacionesActivas { get; set; }
    public string? CorreoNotificacion { get; set; }
    public string? WhatsappContacto { get; set; }
    public bool PermiteChatWhatsapp { get; set; }
    public int MinutosAnticipacionRecordatorio { get; set; }
    public int MinutosToleranciaNoShow { get; set; }
    public string DiasAtencion { get; set; } = string.Empty;
    public string HorarioAtencion { get; set; } = string.Empty;
    public int FechasNoLaborablesCount { get; set; }
    public bool Activo { get; set; }
}

public class EspaciosIndexViewModel : ModuloBaseViewModel
{
    public List<EspacioItemViewModel> Espacios { get; set; } = new();
}

public class EspacioItemViewModel
{
    public int Id { get; set; }
    public string Codigo { get; set; } = string.Empty;
    public string Nombre { get; set; } = string.Empty;
    public string SedeNombre { get; set; } = string.Empty;
    public string TipoDeporteNombre { get; set; } = string.Empty;
    public string TipoSueloNombre { get; set; } = string.Empty;
    public string Estado { get; set; } = string.Empty;
    public string TarifaResumen { get; set; } = string.Empty;
}

public class ReservasIndexViewModel : ModuloBaseViewModel
{
    public List<ReservaItemViewModel> Reservas { get; set; } = new();
    public int TotalReservasListado { get; set; }
    public int TotalPendientesListadoGlobal { get; set; }
    public int TotalPagadasListadoGlobal { get; set; }
    public decimal SaldoTotalListadoGlobal { get; set; }
    public int PaginaListado { get; set; } = 1;
    public int TamanoPaginaListado { get; set; } = 20;
    public int TotalPaginasListado { get; set; } = 1;
    public List<BloqueoHorarioItemViewModel> Bloqueos { get; set; } = new();
    public DateOnly FechaDesde { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly FechaHasta { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(6));
    public DateOnly ListadoFechaDesde { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly ListadoFechaHasta { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(6));
    public int? SedeId { get; set; }
    public int? EspacioDeportivoId { get; set; }
    public int? Estado { get; set; }
    public List<int> EstadosListadoSeleccionados { get; set; } = new();
    public List<SelectListItem> SedesFiltro { get; set; } = new();
    public List<SelectListItem> EspaciosFiltro { get; set; } = new();
    public List<SelectListItem> EstadosFiltro { get; set; } = new();
    public List<SelectListItem> ClientesFiltro { get; set; } = new();
    public List<SelectListItem> TiposDocumentoClientesFiltro { get; set; } = new();
    public List<SelectListItem> FormasPagoFiltro { get; set; } = new();
    public BloqueoHorarioFormViewModel BloqueoForm { get; set; } = new();
    public bool CalendarioUsaHorarioSede { get; set; }
    public bool AtiendeLunes { get; set; } = true;
    public bool AtiendeMartes { get; set; } = true;
    public bool AtiendeMiercoles { get; set; } = true;
    public bool AtiendeJueves { get; set; } = true;
    public bool AtiendeViernes { get; set; } = true;
    public bool AtiendeSabado { get; set; } = true;
    public bool AtiendeDomingo { get; set; } = true;
    public TimeOnly HoraApertura { get; set; } = new(6, 0);
    public TimeOnly HoraCierre { get; set; } = new(23, 0);
    public List<string> FechasNoLaborables { get; set; } = new();
    public int PoliticaConfirmacionPago { get; set; }
    public decimal? PorcentajeAdelantoMinimo { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public string MonedaNombre { get; set; } = "PEN";
}

public class ReservasListadoResumenViewModel
{
    public int TotalPendientes { get; set; }
    public int TotalPagadas { get; set; }
    public decimal SaldoTotal { get; set; }
}

public class ReservaClienteRapidoRequestViewModel
{
    public int NegocioId { get; set; }
    public string TipoDocumento { get; set; } = "0";
    public string? NumeroDocumento { get; set; }
    public string? Nombres { get; set; }
    public string? Apellidos { get; set; }
    public string? RazonSocial { get; set; }
    public string? NombreEquipo { get; set; }
    public string? Telefono { get; set; }
    public string? Correo { get; set; }
}

public class ReservaItemViewModel
{
    public int Id { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public string? Equipo { get; set; }
    public string Espacio { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public decimal Total { get; set; }
    public decimal Adelanto { get; set; }
    public decimal SaldoPendiente { get; set; }
    public string Estado { get; set; } = string.Empty;
}

public class ReservaCalendarioEventoViewModel
{
    public int Id { get; set; }
    public string TipoEvento { get; set; } = string.Empty;
    public string Titulo { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public int? Estado { get; set; }
    public string? EstadoCodigo { get; set; }
    public string? EstadoTexto { get; set; }
    public string? Motivo { get; set; }
    public string? Color { get; set; }
    public decimal TotalReserva { get; set; }
    public int? EspacioDeportivoId { get; set; }
    public string Espacio { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
}

public class ReservaCotizacionViewModel
{
    public bool Ok { get; set; }
    public string Mensaje { get; set; } = string.Empty;
    public decimal PrecioBase { get; set; }
    public decimal DescuentoPct { get; set; }
    public decimal PrecioFinal { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public string MonedaNombre { get; set; } = "PEN";
    public int PoliticaConfirmacionPago { get; set; }
    public decimal? PorcentajeAdelantoMinimo { get; set; }
}

public class BloqueoHorarioItemViewModel
{
    public int Id { get; set; }
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string Motivo { get; set; } = string.Empty;
    public bool Activo { get; set; }
}

public class PagosIndexViewModel : ModuloBaseViewModel
{
    public string? Buscar { get; set; }
    public int Pagina { get; set; } = 1;
    public int TamanoPagina { get; set; } = 20;
    public int TotalRegistros { get; set; }
    public int TotalPaginas { get; set; } = 1;
    public string MonedaSimbolo { get; set; } = "S/";
    public bool EmisionComprobantesElectronicos { get; set; }
    public bool EmisionReciboInterno { get; set; }
    public List<PagoReservaResumenViewModel> Pagos { get; set; } = new();
}

public class PagoReservaResumenViewModel
{
    public int ReservaId { get; set; }
    public string ReservaCodigo { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public string Cliente { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public decimal MontoTotal { get; set; }
    public decimal SaldoPendiente { get; set; }
    public string FormaPagoResumen { get; set; } = string.Empty;
    public int CantidadPagos { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public bool PagadaCompleta { get; set; }
    public bool TieneComprobanteActivo { get; set; }
    public string Referencia { get; set; } = string.Empty;
}

public class PagoReservaEditViewModel : ModuloBaseViewModel
{
    public int ReservaId { get; set; }
    public string ReservaCodigo { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public string Cliente { get; set; } = string.Empty;
    public DateOnly FechaReserva { get; set; }
    public TimeOnly HoraInicioReserva { get; set; }
    public TimeOnly HoraFinReserva { get; set; }
    public decimal TotalReserva { get; set; }
    public decimal TotalPagado { get; set; }
    public decimal SaldoPendiente { get; set; }
    public string MonedaSimbolo { get; set; } = "S/";
    public int PoliticaConfirmacionPago { get; set; }
    public decimal? PorcentajeAdelantoMinimo { get; set; }
    public bool TieneComprobanteActivo { get; set; }
    public string ReferenciaComprobante { get; set; } = string.Empty;
    public List<PagoReservaDetalleItemViewModel> Pagos { get; set; } = new();

    public bool AgregarNuevoPago { get; set; }
    public DateTime? NuevaFechaPago { get; set; }
    public decimal? NuevoMonto { get; set; }
    public int? NuevaFormaPagoId { get; set; }
    public string? NuevoNumeroOperacion { get; set; }
    public string? NuevaObservacion { get; set; }
    public List<SelectListItem> FormasPago { get; set; } = new();
}

public class PagoReservaDetalleItemViewModel
{
    public int PagoId { get; set; }
    public DateTime FechaPago { get; set; }
    public decimal Monto { get; set; }
    public int FormaPagoId { get; set; }
    public string FormaPagoNombre { get; set; } = string.Empty;
    public string? NumeroOperacion { get; set; }
    public string? Observacion { get; set; }
    public bool Eliminar { get; set; }
}

public class ComprobantesIndexViewModel : ModuloBaseViewModel
{
    public string? Buscar { get; set; }
    public string? CodigoDocumento { get; set; }
    public int Pagina { get; set; } = 1;
    public int TamanoPagina { get; set; } = 20;
    public int TotalRegistros { get; set; }
    public int TotalPaginas { get; set; } = 1;
    public List<SelectListItem> TiposDocumentoFiltro { get; set; } = new();
    public List<ComprobanteItemViewModel> Comprobantes { get; set; } = new();
}

public class ComprobanteItemViewModel
{
    public int Id { get; set; }
    public int ReservaId { get; set; }
    public string Tipo { get; set; } = string.Empty;
    public string SerieNumero { get; set; } = string.Empty;
    public DateTime FechaEmision { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public decimal Total { get; set; }
    public string Estado { get; set; } = string.Empty;
    public int EstadoCodigo { get; set; }
    public string CodigoDocumentoComprobante { get; set; } = string.Empty;
    public string Referencia { get; set; } = string.Empty;
    public bool TieneNotasRelacionadas { get; set; }
    public bool EsTributario { get; set; }
    public string? UrlDescargaProveedor { get; set; }
}

public class ClientesIndexViewModel : ModuloBaseViewModel
{
    public string EstadoFiltro { get; set; } = "activos";
    public string? Buscar { get; set; }
    public int Pagina { get; set; } = 1;
    public int TamanoPagina { get; set; } = 20;
    public int TotalRegistros { get; set; }
    public int TotalPaginas { get; set; } = 1;
    public int TotalActivos { get; set; }
    public int TotalInactivos { get; set; }
    public List<ClienteItemViewModel> Clientes { get; set; } = new();
}

public class PromocionesIndexViewModel : ModuloBaseViewModel
{
    public List<PromocionItemViewModel> Promociones { get; set; } = new();
}

public class MaestrosIndexViewModel : ModuloBaseViewModel
{
    public List<MonedaMaestroItemViewModel> Monedas { get; set; } = new();
    public List<SelectListItem> MonedasSuper { get; set; } = new();
    public List<SelectListItem> TiposSueloSuper { get; set; } = new();
    public List<SelectListItem> TiposDeporteSuper { get; set; } = new();
    public List<MaestroCatalogoItemViewModel> TiposSuelo { get; set; } = new();
    public List<MaestroCatalogoItemViewModel> TiposDeporte { get; set; } = new();
    public List<MaestroCatalogoItemViewModel> FormasPago { get; set; } = new();
    public List<SelectListItem> TiposDocumentoComprobanteSuper { get; set; } = new();
    public List<TipoDocumentoComprobanteNegocioItemViewModel> TiposDocumentoComprobante { get; set; } = new();
}

public class MonedaMaestroItemViewModel
{
    public int Id { get; set; }
    public int MonedaSuperId { get; set; }
    public string Codigo { get; set; } = string.Empty;
    public string Nombre { get; set; } = string.Empty;
    public string? Simbolo { get; set; }
    public bool Activo { get; set; }
}

public class MaestroCatalogoItemViewModel
{
    public int Id { get; set; }
    public int? SuperId { get; set; }
    public string? Codigo { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public bool Activo { get; set; }
}

public class TipoDocumentoComprobanteNegocioItemViewModel
{
    public int Id { get; set; }
    public string CodigoSunat { get; set; } = string.Empty;
    public string Nombre { get; set; } = string.Empty;
    public bool Tributario { get; set; }
    public bool HabilitadoSuper { get; set; }
    public bool Activo { get; set; }
}

public class PromocionItemViewModel
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly FechaInicio { get; set; }
    public DateOnly FechaFin { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public decimal PorcentajeDescuento { get; set; }
    public bool Activo { get; set; }
}

public class ClienteItemViewModel
{
    public int Id { get; set; }
    public string NombresORazonSocial { get; set; } = string.Empty;
    public string? NombreEquipo { get; set; }
    public string TipoDocumento { get; set; } = string.Empty;
    public string NumeroDocumento { get; set; } = string.Empty;
    public string? Telefono { get; set; }
    public string? Correo { get; set; }
    public bool Activo { get; set; }
}
