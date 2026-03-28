namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class ReportesIndexViewModel : ModuloBaseViewModel
{
    public DateOnly FechaDesde { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(-6));
    public DateOnly FechaHasta { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public List<ReporteOcupacionItemViewModel> Ocupacion { get; set; } = new();
    public List<ReporteIngresoDiaItemViewModel> Ingresos { get; set; } = new();
}

public class ReporteOcupacionItemViewModel
{
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
