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
