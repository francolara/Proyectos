namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PersonaDetalleDto
{
    public int IdPersona { get; init; }
    public int IdEmpresa { get; init; }
    public string TipoPersona { get; init; } = string.Empty;
    public string TipoDocumento { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string? ApellidoPaterno { get; init; }
    public string? ApellidoMaterno { get; init; }
    public string? Nombres { get; init; }
    public string? RazonSocial { get; init; }
    public string NombreCompleto { get; init; } = string.Empty;
    public string? CorreoElectronico { get; init; }
    public string? Telefono { get; init; }
    public string? Direccion { get; init; }
    public string? CodigoUbigeo { get; init; }
    public string? CodigoDepartamento { get; init; }
    public string? CodigoProvincia { get; init; }
    public bool EsCliente { get; init; }
    public bool EsProveedor { get; init; }
    public bool Estado { get; init; }
}
