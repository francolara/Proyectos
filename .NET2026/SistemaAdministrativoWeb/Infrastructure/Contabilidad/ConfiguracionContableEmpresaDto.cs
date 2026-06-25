namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionContableEmpresaDto
{
    public List<ConfiguracionContableProvisionDto> Provisiones { get; } = [];
    public List<ConfiguracionDocumentoEmpresaDto> Documentos { get; } = [];
    public List<ConfiguracionImpuestoEmpresaDto> Impuestos { get; } = [];
}
