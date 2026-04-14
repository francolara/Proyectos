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
    public List<ReporteOcupacionItemViewModel> Ocupacion { get; set; } = new();
    public List<ReporteIngresoDiaItemViewModel> Ingresos { get; set; } = new();
    public List<ReporteIngresoDiaItemViewModel> IngresosPeriodoAnterior { get; set; } = new();
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
