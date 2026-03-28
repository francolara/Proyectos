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
    public List<NegocioAccesoViewModel> Negocios { get; set; } = new();
    public List<PermisoModuloViewModel> Modulos { get; set; } = new();
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
