using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class SolicitudesIndexViewModel : ModuloBaseViewModel
{
    public DateOnly FechaDesde { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(-7));
    public DateOnly FechaHasta { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public int? Estado { get; set; }
    public List<SolicitudPublicaItemViewModel> Solicitudes { get; set; } = new();
}

public class SolicitudPublicaItemViewModel
{
    public int Id { get; set; }
    public string CodigoSolicitud { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string NombreSolicitante { get; set; } = string.Empty;
    public string Telefono { get; set; } = string.Empty;
    public string? Correo { get; set; }
    public int Estado { get; set; }
    public int? ReservaId { get; set; }
    public DateTime FechaRegistro { get; set; }
}

public class SolicitudEstadoFormViewModel
{
    public int NegocioId { get; set; }
    public int Id { get; set; }
    public int Estado { get; set; }

    [StringLength(300, ErrorMessage = "El campo {0} excede la longitud permitida.")]
    public string? ComentarioGestion { get; set; }
}

public class SolicitudConvertirFormViewModel
{
    public int NegocioId { get; set; }
    public int Id { get; set; }

    [Range(0, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Total { get; set; }

    [Range(0, 999999, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public decimal Adelanto { get; set; }

    [Range(1, 6, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int EstadoReserva { get; set; } = 1;
}
