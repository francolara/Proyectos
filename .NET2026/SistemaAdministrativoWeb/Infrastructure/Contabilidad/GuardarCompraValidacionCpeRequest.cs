namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarCompraValidacionCpeRequest
{
    public int IdCompra { get; init; }
    public int IdEmpresa { get; init; }
    public DateTime FechaValidacionCpe { get; init; }
    public string EstadoValidacionCpe { get; init; } = string.Empty;
    public string? MensajeValidacionCpe { get; init; }
    public string? UsuarioRegistro { get; init; }
}
