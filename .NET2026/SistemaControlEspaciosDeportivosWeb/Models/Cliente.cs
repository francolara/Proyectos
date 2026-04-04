using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.Models;

public class Cliente
{
    public int Id { get; set; }

    [MaxLength(200)]
    public string NombresORazonSocial { get; set; } = string.Empty;

    [MaxLength(20)]
    public string TipoDocumento { get; set; } = string.Empty;

    [MaxLength(20)]
    public string NumeroDocumento { get; set; } = string.Empty;

    [MaxLength(20)]
    public string? Telefono { get; set; }

    [MaxLength(200)]
    public string? Correo { get; set; }

    [MaxLength(250)]
    public string? DireccionFiscal { get; set; }
    [MaxLength(6)]
    public string? CodigoUbigeo { get; set; }

    public bool Activo { get; set; } = true;

    public ICollection<Reserva> Reservas { get; set; } = new List<Reserva>();
    public ICollection<ComprobanteElectronico> ComprobantesElectronicos { get; set; } = new List<ComprobanteElectronico>();
}
