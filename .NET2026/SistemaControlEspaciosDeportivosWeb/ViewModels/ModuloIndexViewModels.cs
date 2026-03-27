using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class ModuloBaseViewModel
{
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;
    public string ModuloCodigo { get; set; } = string.Empty;
    public string ModuloNombre { get; set; } = string.Empty;
    public bool PuedeCrear { get; set; }
    public bool PuedeEditar { get; set; }
    public bool PuedeEliminar { get; set; }
    public string? Mensaje { get; set; }
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
    public string Estado { get; set; } = string.Empty;
}

public class ReservasIndexViewModel : ModuloBaseViewModel
{
    public List<ReservaItemViewModel> Reservas { get; set; } = new();
    public List<BloqueoHorarioItemViewModel> Bloqueos { get; set; } = new();
    public DateOnly FechaDesde { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly FechaHasta { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(6));
    public int? SedeId { get; set; }
    public int? EspacioDeportivoId { get; set; }
    public int? Estado { get; set; }
    public List<SelectListItem> SedesFiltro { get; set; } = new();
    public List<SelectListItem> EspaciosFiltro { get; set; } = new();
    public List<SelectListItem> EstadosFiltro { get; set; } = new();
    public BloqueoHorarioFormViewModel BloqueoForm { get; set; } = new();
}

public class ReservaItemViewModel
{
    public int Id { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public decimal Total { get; set; }
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
    public string? Color { get; set; }
    public int? EspacioDeportivoId { get; set; }
    public string Espacio { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
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
    public List<PagoItemViewModel> Pagos { get; set; } = new();
}

public class PagoItemViewModel
{
    public int Id { get; set; }
    public int ReservaId { get; set; }
    public DateTime FechaPago { get; set; }
    public decimal Monto { get; set; }
    public string FormaPago { get; set; } = string.Empty;
}

public class ComprobantesIndexViewModel : ModuloBaseViewModel
{
    public List<ComprobanteItemViewModel> Comprobantes { get; set; } = new();
}

public class ComprobanteItemViewModel
{
    public int Id { get; set; }
    public string Tipo { get; set; } = string.Empty;
    public string SerieNumero { get; set; } = string.Empty;
    public DateTime FechaEmision { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public decimal Total { get; set; }
    public string Estado { get; set; } = string.Empty;
}

public class ClientesIndexViewModel : ModuloBaseViewModel
{
    public List<ClienteItemViewModel> Clientes { get; set; } = new();
}

public class PromocionesIndexViewModel : ModuloBaseViewModel
{
    public List<PromocionItemViewModel> Promociones { get; set; } = new();
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
    public string TipoDocumento { get; set; } = string.Empty;
    public string NumeroDocumento { get; set; } = string.Empty;
    public string? Telefono { get; set; }
    public string? Correo { get; set; }
    public bool Activo { get; set; }
}
