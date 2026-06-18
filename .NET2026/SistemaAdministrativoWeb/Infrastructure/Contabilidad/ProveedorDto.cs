namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ProveedorDto
{
    public int IdProveedor { get; init; }
    public int IdEmpresa { get; init; }
    public int IdPersona { get; init; }
    public string CodigoProveedor { get; init; } = string.Empty;
    public string TipoDocumento { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string NombreCompleto { get; init; } = string.Empty;
    public string? CorreoElectronico { get; init; }
    public string? Telefono { get; init; }
    public string? Contacto { get; init; }
    public string? CuentaDetraccion { get; init; }
    public string? Observacion { get; init; }
    public bool Estado { get; init; }
}
