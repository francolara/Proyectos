namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PeriodoContableEstadoDto
{
    public int? IdPeriodoContableEstado { get; init; }
    public int IdEmpresa { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public bool Cerrado { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string? UsuarioRegistro { get; init; }
    public DateTime? FechaCierre { get; init; }
    public string? UsuarioCierre { get; init; }
    public DateTime? FechaApertura { get; init; }
    public string? UsuarioApertura { get; init; }
}
