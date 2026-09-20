namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class OrigenDto
{
    public int IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string ModuloOrigen { get; init; } = string.Empty;
    public bool PermiteRegistroManual { get; init; }
    public bool Estado { get; init; }
    public string? UsuarioCreacion { get; init; }
    public DateTime? FechaRegistro { get; init; }
    public string? UsuarioActualizacion { get; init; }
    public DateTime? FechaActualizacion { get; init; }
}
