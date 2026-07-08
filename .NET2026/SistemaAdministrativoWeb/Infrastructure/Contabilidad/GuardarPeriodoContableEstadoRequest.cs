namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarPeriodoContableEstadoRequest
{
    public int IdEmpresa { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public bool Cerrado { get; init; }
    public string? UsuarioRegistro { get; init; }
}
