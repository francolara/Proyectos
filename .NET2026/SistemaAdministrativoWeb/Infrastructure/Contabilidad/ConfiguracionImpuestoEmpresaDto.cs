namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionImpuestoEmpresaDto
{
    public int IdTipoImpuesto { get; init; }
    public string CodigoSunat { get; init; } = string.Empty;
    public string NombreImpuesto { get; init; } = string.Empty;
    public int? IdPlanCuenta { get; init; }
    public string CuentaTexto { get; init; } = string.Empty;
    public bool Activo { get; init; } = true;
}
