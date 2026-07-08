namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MigoRucResultDto
{
    public string Ruc { get; init; } = string.Empty;
    public string NombreORazonSocial { get; init; } = string.Empty;
    public string? Ubigeo { get; init; }
    public string? DireccionSimple { get; init; }
    public string? Direccion { get; init; }
    public string? Distrito { get; init; }
    public string? Provincia { get; init; }
    public string? Departamento { get; init; }
    public string? EstadoContribuyente { get; init; }
    public string? CondicionDomicilio { get; init; }
}
