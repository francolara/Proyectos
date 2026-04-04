namespace SistemaControlEspaciosDeportivosWeb.Models;

public class Negocio
{
    public int Id { get; set; }
    public string NombreComercial { get; set; } = string.Empty;
    public string? RazonSocial { get; set; }
    public string? DocumentoFiscal { get; set; }
    public string? TipoDocumentoFiscal { get; set; }
    public string? NumeroDocumentoFiscal { get; set; }
    public string? DireccionFiscal { get; set; }
    public string? CodigoUbigeo { get; set; }
    public int? MonedaId { get; set; }
    public bool Activo { get; set; } = true;
    public DateTime FechaRegistro { get; set; } = DateTime.UtcNow;

    public ICollection<Sede> Sedes { get; set; } = new List<Sede>();
    public ICollection<UsuarioNegocio> UsuariosNegocio { get; set; } = new List<UsuarioNegocio>();
    public ICollection<ComprobanteElectronico> ComprobantesElectronicos { get; set; } = new List<ComprobanteElectronico>();
}
