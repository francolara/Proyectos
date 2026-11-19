using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using System.Security.Claims;

namespace SistemaAdministrativoWeb.ViewComponents;

public sealed class ReporteEmpresaViewComponent(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IEmpresaRepository empresaRepository) : ViewComponent
{
    public async Task<IViewComponentResult> InvokeAsync(string empresaNombre, string? rucEmpresa)
    {
        if (rucEmpresa is null &&
            currentCompanyAccessor.EmpresaId is int empresaId &&
            HttpContext.User.FindFirstValue(ClaimTypes.NameIdentifier) is { Length: > 0 } usuarioId)
        {
            var empresas = await empresaRepository.ListarPorUsuarioAsync(usuarioId, HttpContext.RequestAborted);
            var empresa = empresas.FirstOrDefault(item => item.IdEmpresa == empresaId);
            empresaNombre = empresa?.RazonSocial ?? empresaNombre;
            rucEmpresa = empresa?.Ruc;
        }

        return View((empresaNombre, rucEmpresa));
    }
}
