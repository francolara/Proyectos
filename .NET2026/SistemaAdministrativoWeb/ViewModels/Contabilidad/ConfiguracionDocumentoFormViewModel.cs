namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionDocumentoFormViewModel
{
    public int IdTipoComprobante { get; set; }
    public string CodigoTipoComprobante { get; set; } = string.Empty;
    public string Descripcion { get; set; } = string.Empty;
    public bool UsoCompras { get; set; }
    public bool UsoVentas { get; set; }
    public int? IdCuentaVentaSoles { get; set; }
    public string CuentaVentaSolesTexto { get; set; } = string.Empty;
    public int? IdCuentaVentaDolares { get; set; }
    public string CuentaVentaDolaresTexto { get; set; } = string.Empty;
    public int? IdCuentaCompraSoles { get; set; }
    public string CuentaCompraSolesTexto { get; set; } = string.Empty;
    public int? IdCuentaCompraDolares { get; set; }
    public string CuentaCompraDolaresTexto { get; set; } = string.Empty;
    public bool Activo { get; set; } = true;
}
