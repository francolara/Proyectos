using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.Models;

public class ComprobanteElectronico
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public int ReservaId { get; set; }
    public int ClienteId { get; set; }
    public TipoComprobante TipoComprobante { get; set; }

    [MaxLength(4)]
    public string Serie { get; set; } = string.Empty;

    public int Numero { get; set; }
    public DateTime FechaEmision { get; set; } = DateTime.UtcNow;
    public TipoMoneda TipoMoneda { get; set; } = TipoMoneda.PEN;

    [MaxLength(4)]
    public string CodigoTipoOperacionSunat { get; set; } = "0101";

    [MaxLength(4)]
    public string CodigoTipoDocumentoClienteSunat { get; set; } = "1";

    [MaxLength(8)]
    public string? CodigoHashCpe { get; set; }

    [MaxLength(40)]
    public string? NumeroTicketSunat { get; set; }

    [MaxLength(50)]
    public string? CodigoRespuestaSunat { get; set; }

    [MaxLength(500)]
    public string? MensajeRespuestaSunat { get; set; }

    public decimal SubTotal { get; set; }
    public decimal Igv { get; set; }
    public decimal Total { get; set; }
    public EstadoComprobanteElectronico Estado { get; set; } = EstadoComprobanteElectronico.PendienteEnvio;
    public DateTime FechaRegistro { get; set; } = DateTime.UtcNow;
    public DateTime? FechaActualizacion { get; set; }
    public string? UsuarioCreacion { get; set; }
    public string? UsuarioActualizacion { get; set; }

    public Negocio? Negocio { get; set; }
    public Reserva? Reserva { get; set; }
    public Cliente? Cliente { get; set; }
    public ICollection<ComprobanteDetalle> Detalles { get; set; } = new List<ComprobanteDetalle>();
}
