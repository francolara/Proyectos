using System.Security.Claims;

namespace SistemaAdministrativoWeb.Infrastructure.Empresas;

public sealed class ActiveCompanySessionValidationMiddleware(RequestDelegate next)
{
    private const string EmpresaValidationKey = "EmpresaActivaValidacion";

    public async Task InvokeAsync(HttpContext context, ICurrentCompanyAccessor currentCompanyAccessor, IEmpresaRepository empresaRepository)
    {
        if (context.User.Identity?.IsAuthenticated == true
            && !context.User.IsInRole("SuperAdmin")
            && currentCompanyAccessor.TieneEmpresaActiva
            && currentCompanyAccessor.EmpresaId.HasValue)
        {
            var aspNetUserId = context.User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrWhiteSpace(aspNetUserId))
            {
                currentCompanyAccessor.LimpiarEmpresa();
            }
            else
            {
                var validationToken = $"{aspNetUserId}:{currentCompanyAccessor.EmpresaId.Value}";
                var validationState = context.Session.GetString(EmpresaValidationKey);

                if (!string.Equals(validationState, validationToken, StringComparison.Ordinal))
                {
                    var empresas = await empresaRepository.ListarPorUsuarioAsync(aspNetUserId, context.RequestAborted);
                    var empresaActiva = empresas.FirstOrDefault(x => x.IdEmpresa == currentCompanyAccessor.EmpresaId.Value);

                    if (empresaActiva is null)
                    {
                        currentCompanyAccessor.LimpiarEmpresa();
                    }
                    else
                    {
                        currentCompanyAccessor.EstablecerEmpresa(empresaActiva.IdEmpresa, empresaActiva.RazonSocial);
                        context.Session.SetString(EmpresaValidationKey, validationToken);
                    }
                }
            }
        }

        await next(context);
    }
}
