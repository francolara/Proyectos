using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.Models;

public class BitacoraAuditoria
{
    public long Id { get; set; }
    public int? NegocioId { get; set; }

    [MaxLength(50)]
    public string Modulo { get; set; } = string.Empty;

    [MaxLength(20)]
    public string Accion { get; set; } = string.Empty;

    [MaxLength(80)]
    public string Entidad { get; set; } = string.Empty;

    [MaxLength(80)]
    public string EntidadId { get; set; } = string.Empty;

    [MaxLength(450)]
    public string UsuarioId { get; set; } = string.Empty;

    [MaxLength(200)]
    public string? UsuarioNombre { get; set; }

    [MaxLength(200)]
    public string? UsuarioCorreo { get; set; }

    [MaxLength(4000)]
    public string? DetalleJson { get; set; }

    public DateTime FechaRegistro { get; set; } = DateTime.UtcNow;
}

