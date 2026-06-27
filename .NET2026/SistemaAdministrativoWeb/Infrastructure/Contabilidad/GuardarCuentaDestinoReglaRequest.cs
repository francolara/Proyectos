namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarCuentaDestinoReglaRequest
{
    public int IdEmpresa { get; init; }
    public int IdPlanCuentaOrigen { get; init; }
    public bool Activo { get; init; }
    public string? Observacion { get; init; }
    public string? UsuarioRegistro { get; init; }
    public IReadOnlyCollection<GuardarCuentaDestinoReglaDetalleRequest> Detalles { get; init; } = [];
}
