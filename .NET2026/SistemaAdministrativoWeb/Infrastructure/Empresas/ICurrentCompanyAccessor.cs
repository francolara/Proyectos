namespace SistemaAdministrativoWeb.Infrastructure.Empresas;

public interface ICurrentCompanyAccessor
{
    int? EmpresaId { get; }
    string? EmpresaNombre { get; }
    bool TieneEmpresaActiva { get; }
    void EstablecerEmpresa(int empresaId, string empresaNombre);
    void LimpiarEmpresa();
}
