namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarCuentaDestinoReglaDetalleRequest
{
    public short Orden { get; init; }
    public int IdPlanCuentaDestinoCargo { get; init; }
    public int IdPlanCuentaDestinoAbono { get; init; }
    public decimal Porcentaje { get; init; }
    public bool Activo { get; init; }
}
