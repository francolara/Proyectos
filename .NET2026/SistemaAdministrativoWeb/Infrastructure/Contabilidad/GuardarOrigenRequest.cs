namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarOrigenRequest
{
    public int? IdOrigen { get; init; }
    public int IdEmpresa { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string ModuloOrigen { get; init; } = string.Empty;
    public bool PermiteRegistroManual { get; init; }
    public bool Estado { get; init; } = true;
    public string? UsuarioRegistro { get; init; }
}
