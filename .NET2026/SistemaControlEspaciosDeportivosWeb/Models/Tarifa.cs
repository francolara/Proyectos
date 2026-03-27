namespace SistemaControlEspaciosDeportivosWeb.Models;

public class Tarifa
{
    public int Id { get; set; }
    public int EspacioDeportivoId { get; set; }
    public DayOfWeek DiaSemana { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public decimal Precio { get; set; }
    public bool Activa { get; set; } = true;

    public EspacioDeportivo? EspacioDeportivo { get; set; }
}

