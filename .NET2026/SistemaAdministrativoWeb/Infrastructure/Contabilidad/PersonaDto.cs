namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PersonaDto
{
    public int IdPersona { get; init; }
    public int IdEmpresa { get; init; }
    public string TipoPersona { get; init; } = string.Empty;
    public string TipoDocumento { get; init; } = string.Empty;
    public string NombreTipoDocumento { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string NombreCompleto { get; init; } = string.Empty;
    public string? CorreoElectronico { get; init; }
    public string? Telefono { get; init; }
    public string? Direccion { get; init; }
    public string? CodigoUbigeo { get; init; }
    public string? Departamento { get; init; }
    public string? Provincia { get; init; }
    public string? Distrito { get; init; }
    public bool EsCliente { get; init; }
    public bool EsProveedor { get; init; }
    public bool Estado { get; init; }
}
