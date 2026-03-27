using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Models;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class SedeFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required]
    [StringLength(150)]
    public string Nombre { get; set; } = string.Empty;

    [Required]
    [StringLength(250)]
    public string Direccion { get; set; } = string.Empty;

    [StringLength(20)]
    public string? Telefono { get; set; }

    public bool Activo { get; set; } = true;

    public List<int> ServiciosSeleccionados { get; set; } = new();
    public List<SelectListItem> ServiciosDisponibles { get; set; } = new();

    public bool NotificacionesActivas { get; set; } = true;

    [Range(5, 1440)]
    public int MinutosAnticipacionRecordatorio { get; set; } = 90;

    [Range(0, 240)]
    public int MinutosToleranciaNoShow { get; set; } = 30;

    [StringLength(200)]
    [EmailAddress]
    public string? CorreoNotificacion { get; set; }

    [StringLength(20)]
    public string? WhatsappContacto { get; set; }

    public bool PermiteChatWhatsapp { get; set; }
}

public class EspacioFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required]
    public int SedeId { get; set; }

    [Required]
    public int TipoDeporteId { get; set; }

    [Required]
    [StringLength(20)]
    public string Codigo { get; set; } = string.Empty;

    [Required]
    [StringLength(150)]
    public string Nombre { get; set; } = string.Empty;

    [Range(1, 200)]
    public int Capacidad { get; set; } = 10;

    public bool TieneIluminacion { get; set; }
    public bool Techada { get; set; }
    public EstadoEspacioDeportivo Estado { get; set; } = EstadoEspacioDeportivo.Activo;

    public List<SelectListItem> Sedes { get; set; } = new();
    public List<SelectListItem> TiposDeporte { get; set; } = new();
}

public class ReservaFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required]
    public int EspacioDeportivoId { get; set; }

    [Required]
    public int ClienteId { get; set; }

    [Required]
    public DateOnly Fecha { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required]
    public TimeOnly HoraInicio { get; set; } = new(18, 0);

    [Required]
    public TimeOnly HoraFin { get; set; } = new(19, 0);

    [Range(0, 999999)]
    public decimal Total { get; set; }

    [Range(0, 999999)]
    public decimal Adelanto { get; set; }

    public EstadoReserva Estado { get; set; } = EstadoReserva.Pendiente;

    public List<SelectListItem> Espacios { get; set; } = new();
    public List<SelectListItem> Clientes { get; set; } = new();
}

public class BloqueoHorarioFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }

    [Required]
    public int EspacioDeportivoId { get; set; }

    [Required]
    public DateOnly Fecha { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required]
    public TimeOnly HoraInicio { get; set; } = new(8, 0);

    [Required]
    public TimeOnly HoraFin { get; set; } = new(9, 0);

    [Required]
    [StringLength(250)]
    public string Motivo { get; set; } = string.Empty;

    public List<SelectListItem> Espacios { get; set; } = new();
}

public class PagoFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required]
    public int ReservaId { get; set; }

    [Range(0.01, 999999)]
    public decimal Monto { get; set; }

    public DateTime FechaPago { get; set; } = DateTime.Now;
    public FormaPago FormaPago { get; set; } = FormaPago.Efectivo;

    [StringLength(50)]
    public string? NumeroOperacion { get; set; }

    [StringLength(300)]
    public string? Observacion { get; set; }

    public List<SelectListItem> Reservas { get; set; } = new();
}

public class ComprobanteFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    [Required]
    public int ReservaId { get; set; }

    [Required]
    public TipoComprobante TipoComprobante { get; set; } = TipoComprobante.Boleta;

    [Required]
    [StringLength(4)]
    public string Serie { get; set; } = "B001";

    [Range(1, int.MaxValue)]
    public int Numero { get; set; }

    public DateTime FechaEmision { get; set; } = DateTime.Now;
    public TipoMoneda TipoMoneda { get; set; } = TipoMoneda.PEN;

    [Range(0, 999999)]
    public decimal SubTotal { get; set; }

    [Range(0, 999999)]
    public decimal Igv { get; set; }

    [Range(0, 999999)]
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

    [Required]
    [StringLength(200)]
    public string NombresORazonSocial { get; set; } = string.Empty;

    [Required]
    [StringLength(20)]
    public string TipoDocumento { get; set; } = "DNI";

    [Required]
    [StringLength(20)]
    public string NumeroDocumento { get; set; } = string.Empty;

    [StringLength(20)]
    public string? Telefono { get; set; }

    [StringLength(200)]
    [EmailAddress]
    public string? Correo { get; set; }

    [StringLength(250)]
    public string? DireccionFiscal { get; set; }

    public bool Activo { get; set; } = true;
}

public class PromocionFormViewModel
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;

    public int? SedeId { get; set; }
    public int? EspacioDeportivoId { get; set; }

    [Required]
    [StringLength(150)]
    public string Nombre { get; set; } = string.Empty;

    [Required]
    public DateOnly FechaInicio { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required]
    public DateOnly FechaFin { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required]
    public TimeOnly HoraInicio { get; set; } = new(8, 0);

    [Required]
    public TimeOnly HoraFin { get; set; } = new(10, 0);

    [Range(0, 100)]
    public decimal PorcentajeDescuento { get; set; }

    public bool Activo { get; set; } = true;

    public List<SelectListItem> Sedes { get; set; } = new();
    public List<SelectListItem> Espacios { get; set; } = new();
}
