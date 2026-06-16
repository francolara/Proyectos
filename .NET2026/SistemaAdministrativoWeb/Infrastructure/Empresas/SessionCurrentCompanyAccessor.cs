namespace SistemaAdministrativoWeb.Infrastructure.Empresas;

public sealed class SessionCurrentCompanyAccessor(IHttpContextAccessor httpContextAccessor) : ICurrentCompanyAccessor
{
    private const string EmpresaIdKey = "EmpresaActivaId";
    private const string EmpresaNombreKey = "EmpresaActivaNombre";

    public int? EmpresaId => httpContextAccessor.HttpContext?.Session.GetInt32(EmpresaIdKey);

    public string? EmpresaNombre => httpContextAccessor.HttpContext?.Session.GetString(EmpresaNombreKey);

    public bool TieneEmpresaActiva => EmpresaId.HasValue;

    public void EstablecerEmpresa(int empresaId, string empresaNombre)
    {
        var session = httpContextAccessor.HttpContext?.Session;
        if (session is null)
        {
            return;
        }

        session.SetInt32(EmpresaIdKey, empresaId);
        session.SetString(EmpresaNombreKey, empresaNombre);
    }

    public void LimpiarEmpresa()
    {
        var session = httpContextAccessor.HttpContext?.Session;
        if (session is null)
        {
            return;
        }

        session.Remove(EmpresaIdKey);
        session.Remove(EmpresaNombreKey);
    }
}
