namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class ReservaMoverRequestViewModel
{
    public int NegocioId { get; set; }
    public int ReservaId { get; set; }
    public DateTime Inicio { get; set; }
    public DateTime Fin { get; set; }
}

public class ReservaEstadoRapidoRequestViewModel
{
    public int NegocioId { get; set; }
    public int ReservaId { get; set; }
    public int NuevoEstado { get; set; }
}

public class ReservaDisponibilidadValidacionViewModel
{
    public bool Disponible { get; set; }
    public string Mensaje { get; set; } = string.Empty;
    public string? ConflictoTipo { get; set; }
    public int? ConflictoId { get; set; }
}
