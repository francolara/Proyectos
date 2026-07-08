namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GenerarAjusteCuentaProcesoRequest
{
    public int IdEmpresa { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public string? UsuarioRegistro { get; init; }
}
