using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class ReportesIndexViewModel : ModuloBaseViewModel
{
    public string Preset { get; set; } = "7d";
    public DateOnly FechaDesde { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(-6));
    public DateOnly FechaHasta { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly FechaDesdeAnterior { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(-13));
    public DateOnly FechaHastaAnterior { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(-7));
    public int DiasPeriodo { get; set; } = 7;
    public int? SedeId { get; set; }
    public List<SelectListItem> SedesFiltro { get; set; } = new();
    public ReporteResumenOperativoViewModel ResumenActual { get; set; } = new();
    public ReporteResumenOperativoViewModel ResumenAnterior { get; set; } = new();
    public ReporteResumenCobranzaViewModel CobranzaActual { get; set; } = new();
    public ReporteResumenCobranzaViewModel CobranzaAnterior { get; set; } = new();
    public List<ReporteOcupacionItemViewModel> Ocupacion { get; set; } = new();
    public List<ReporteReservaDiaItemViewModel> ReservasPorDia { get; set; } = new();
    public List<ReporteReservaDiaItemViewModel> ReservasPeriodoAnterior { get; set; } = new();
    public List<ReporteIngresoDiaItemViewModel> Ingresos { get; set; } = new();
    public List<ReporteIngresoDiaItemViewModel> IngresosPeriodoAnterior { get; set; } = new();
    public List<ReportePagoDetalleItemViewModel> DetallePagos { get; set; } = new();
    public List<ReporteReservaDetalleItemViewModel> DetalleReservas { get; set; } = new();
}

public class ReportePagoDetalleItemViewModel
{
    public int PagoId { get; set; }
    public DateTime FechaPago { get; set; }
    public int ReservaId { get; set; }
    public DateOnly FechaReserva { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public string ClienteDocumento { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public string FormaPago { get; set; } = string.Empty;
    public string? NumeroOperacion { get; set; }
    public decimal Monto { get; set; }
}

public class ReporteReservaDetalleItemViewModel
{
    public int ReservaId { get; set; }
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public string ClienteDocumento { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public int EstadoCodigo { get; set; }
    public string Estado { get; set; } = string.Empty;
    public string CanalOrigen { get; set; } = string.Empty;
    public decimal Total { get; set; }
    public decimal Descuento { get; set; }
    public decimal MontoPagado { get; set; }
    public decimal SaldoPendiente { get; set; }
    public string? CodigoCupon { get; set; }
}

public class ReporteOcupacionItemViewModel
{
    public int SedeId { get; set; }
    public int EspacioDeportivoId { get; set; }
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public int CantidadReservas { get; set; }
    public decimal HorasReservadas { get; set; }
    public decimal MontoReservado { get; set; }
    public decimal MontoCobrado { get; set; }
}

public class ReporteIngresoDiaItemViewModel
{
    public DateOnly Fecha { get; set; }
    public int CantidadReservas { get; set; }
    public decimal Ingresos { get; set; }
}

public class ReporteReservaDiaItemViewModel
{
    public DateOnly Fecha { get; set; }
    public int CantidadReservas { get; set; }
    public decimal MontoReservado { get; set; }
}

public class ReporteResumenOperativoViewModel
{
    public int TotalReservas { get; set; }
    public int TotalPendientes { get; set; }
    public int TotalConfirmadas { get; set; }
    public int TotalPagadas { get; set; }
    public int TotalCanceladas { get; set; }
    public int TotalNoShow { get; set; }
    public decimal MontoReservado { get; set; }
    public decimal MontoCobrado { get; set; }
    public decimal SaldoPendiente { get; set; }
}

public class ReporteResumenCobranzaViewModel
{
    public int CantidadPagos { get; set; }
    public int ReservasCobradas { get; set; }
    public decimal MontoCobrado { get; set; }
}
