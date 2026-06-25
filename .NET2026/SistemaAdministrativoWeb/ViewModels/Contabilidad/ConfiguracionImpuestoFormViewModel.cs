namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionImpuestoFormViewModel
{
    public int IdTipoImpuesto { get; set; }
    public string CodigoSunat { get; set; } = string.Empty;
    public string NombreImpuesto { get; set; } = string.Empty;
    public int? IdPlanCuenta { get; set; }
    public string CuentaTexto { get; set; } = string.Empty;
    public bool Activo { get; set; } = true;
}
