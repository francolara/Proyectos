namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class UbigeoDistritoDto
{
    public string CodigoUbigeo { get; init; } = string.Empty;
    public string CodigoDepartamento { get; init; } = string.Empty;
    public string CodigoProvincia { get; init; } = string.Empty;
    public string Nombre { get; init; } = string.Empty;
    public string? Zona { get; init; }
}
