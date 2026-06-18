namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ClienteDto
{
    public int IdCliente { get; init; }
    public int IdEmpresa { get; init; }
    public int IdPersona { get; init; }
    public string CodigoCliente { get; init; } = string.Empty;
    public string TipoDocumento { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string NombreCompleto { get; init; } = string.Empty;
    public string? CorreoElectronico { get; init; }
    public string? Telefono { get; init; }
    public decimal LimiteCredito { get; init; }
    public int DiasCredito { get; init; }
    public string? Observacion { get; init; }
    public bool Estado { get; init; }
}
