namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionDocumentoEmpresaDto
{
    public int IdTipoComprobante { get; init; }
    public string CodigoTipoComprobante { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public bool UsoCompras { get; init; }
    public bool UsoVentas { get; init; }
    public int? IdCuentaVentaSoles { get; init; }
    public string CuentaVentaSolesTexto { get; init; } = string.Empty;
    public int? IdCuentaVentaDolares { get; init; }
    public string CuentaVentaDolaresTexto { get; init; } = string.Empty;
    public int? IdCuentaCompraSoles { get; init; }
    public string CuentaCompraSolesTexto { get; init; } = string.Empty;
    public int? IdCuentaCompraDolares { get; init; }
    public string CuentaCompraDolaresTexto { get; init; } = string.Empty;
    public bool Activo { get; init; } = true;
}
