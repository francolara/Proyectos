namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MigoCpeResultDto
{
    public bool Success { get; init; }
    public string? EstadoComprobante { get; init; }
    public string? EstadoContribuyente { get; init; }
    public string? CondicionDomicilio { get; init; }
    public string? Observaciones { get; init; }
}
