namespace SistemaControlEspaciosDeportivosWeb.Models;

public class Desafio
{
    public int Id { get; set; }
    public string IdUsuarioRetador { get; set; } = string.Empty;
    public string IdUsuarioRetado { get; set; } = string.Empty;
    public int IdDeporte { get; set; }
    public int IdNivel { get; set; }
    public string Distrito { get; set; } = string.Empty;
    public DateOnly FechaTentativa { get; set; }
    public TimeOnly HoraTentativa { get; set; }
    public string? CanchaSugerida { get; set; }
    public string Modalidad { get; set; } = string.Empty;
    public string? Mensaje { get; set; }
    public string FormaPago { get; set; } = string.Empty;
    public string Estado { get; set; } = DesafioEstados.Pendiente;
    public DateTime FechaCreacion { get; set; }
    public DateTime? FechaRespuesta { get; set; }
    public bool Activo { get; set; }
}

public static class DesafioEstados
{
    public const string Pendiente = "Pendiente";
    public const string Aceptado = "Aceptado";
    public const string Rechazado = "Rechazado";
    public const string Cancelado = "Cancelado";
    public const string Finalizado = "Finalizado";
}
