namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CentroCostoDto
{
    public int IdCentroCosto { get; init; }
    public int IdEmpresa { get; init; }
    public string CodigoCentroCosto { get; init; } = string.Empty;
    public string NombreCentroCosto { get; init; } = string.Empty;
    public bool Estado { get; init; }
    public string? UsuarioCreacion { get; init; }
    public DateTime? FechaRegistro { get; init; }
    public string? UsuarioActualizacion { get; init; }
    public DateTime? FechaActualizacion { get; init; }
}
