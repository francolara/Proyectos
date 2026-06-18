namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CuentaDestinoReglaDetalleDto
{
    public int IdCuentaDestinoReglaDetalle { get; init; }
    public int IdCuentaDestinoRegla { get; init; }
    public short Orden { get; init; }
    public int IdPlanCuentaDestinoCargo { get; init; }
    public string CodigoCuentaDestinoCargo { get; init; } = string.Empty;
    public string NombreCuentaDestinoCargo { get; init; } = string.Empty;
    public int IdPlanCuentaDestinoAbono { get; init; }
    public string CodigoCuentaDestinoAbono { get; init; } = string.Empty;
    public string NombreCuentaDestinoAbono { get; init; } = string.Empty;
    public decimal Porcentaje { get; init; }
    public bool Activo { get; init; }
}
