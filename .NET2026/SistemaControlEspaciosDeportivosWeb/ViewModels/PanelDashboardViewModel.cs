using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class PanelDashboardViewModel
{
    public string? Mensaje { get; set; }
    public int NegocioSeleccionadoId { get; set; }
    public string RolActual { get; set; } = string.Empty;
    public int TotalSedes { get; set; }
    public int TotalEspacios { get; set; }
    public int ReservasHoy { get; set; }
    public decimal IngresosHoy { get; set; }
    public decimal OcupacionHoyPct { get; set; }
    public int NoShowMes { get; set; }
    public decimal TicketPromedioMes { get; set; }
    public int? SedeId { get; set; }
    public List<SelectListItem> SedesFiltro { get; set; } = new();
    public DateOnly FechaDesde { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(-6));
    public DateOnly FechaHasta { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public string MonedaSimbolo { get; set; } = "S/";

    public decimal IngresosPeriodo { get; set; }
    public int ReservasPeriodo { get; set; }
    public decimal TicketPromedioPeriodo { get; set; }
    public decimal IngresosPeriodoAnterior { get; set; }
    public int ReservasPeriodoAnterior { get; set; }
    public decimal TicketPromedioPeriodoAnterior { get; set; }
    public decimal OcupacionDiaAnteriorPct { get; set; }

    public List<DashboardSerieItemViewModel> SerieIngresosPorDia { get; set; } = new();
    public List<DashboardSerieItemViewModel> SerieReservasPorDia { get; set; } = new();
    public List<DashboardTopEspacioViewModel> TopEspacios { get; set; } = new();
    public List<DashboardReservaAccionViewModel> ReservasPendientesConfirmacion { get; set; } = new();
    public List<DashboardReservaAccionViewModel> ReservasConSaldoPendiente { get; set; } = new();
    public List<NegocioAccesoViewModel> Negocios { get; set; } = new();
    public List<PermisoModuloViewModel> Modulos { get; set; } = new();
}

public class DashboardSerieItemViewModel
{
    public DateOnly Fecha { get; set; }
    public decimal Valor { get; set; }
}

public class DashboardTopEspacioViewModel
{
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public int Reservas { get; set; }
    public decimal Horas { get; set; }
    public decimal Cobrado { get; set; }
}

public class DashboardReservaAccionViewModel
{
    public int ReservaId { get; set; }
    public string ReservaCodigo { get; set; } = string.Empty;
    public string Cliente { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public decimal MontoTotal { get; set; }
    public decimal SaldoPendiente { get; set; }
    public string Estado { get; set; } = string.Empty;
}

public class NegocioAccesoViewModel
{
    public int NegocioId { get; set; }
    public string NombreNegocio { get; set; } = string.Empty;
    public string Rol { get; set; } = string.Empty;
}

public class PermisoModuloViewModel
{
    public int ModuloSistemaId { get; set; }
    public string ModuloCodigo { get; set; } = string.Empty;
    public string ModuloNombre { get; set; } = string.Empty;
    public bool PuedeVer { get; set; }
    public bool PuedeCrear { get; set; }
    public bool PuedeEditar { get; set; }
    public bool PuedeEliminar { get; set; }
}
