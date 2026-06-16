using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Empresas;
using System.Security.Claims;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class EmpresaContextoController(
    IEmpresaRepository empresaRepository,
    ICurrentCompanyAccessor currentCompanyAccessor) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        var model = new SeleccionEmpresaViewModel
        {
            Empresas = await ObtenerEmpresasAsync(cancellationToken)
        };

        if (model.Empresas.Count == 1 && !currentCompanyAccessor.TieneEmpresaActiva)
        {
            var unica = model.Empresas[0];
            currentCompanyAccessor.EstablecerEmpresa(unica.IdEmpresa, unica.RazonSocial);
            return RedirectToAction("Index", "Home");
        }

        if (currentCompanyAccessor.TieneEmpresaActiva)
        {
            model.IdEmpresaSeleccionada = currentCompanyAccessor.EmpresaId ?? 0;
        }
        else
        {
            var predeterminada = model.Empresas.FirstOrDefault(x => x.EsEmpresaPredeterminada);
            if (predeterminada is not null)
            {
                model.IdEmpresaSeleccionada = predeterminada.IdEmpresa;
            }
        }

        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Seleccionar(SeleccionEmpresaViewModel model, CancellationToken cancellationToken)
    {
        var empresas = await ObtenerEmpresasAsync(cancellationToken);
        var empresa = empresas.FirstOrDefault(x => x.IdEmpresa == model.IdEmpresaSeleccionada);

        if (empresa is null)
        {
            ModelState.AddModelError(nameof(model.IdEmpresaSeleccionada), "Seleccione una empresa valida.");
            model.Empresas = empresas;
            return View("Index", model);
        }

        currentCompanyAccessor.EstablecerEmpresa(empresa.IdEmpresa, empresa.RazonSocial);
        return RedirectToAction("Index", "Home");
    }

    private async Task<List<EmpresaDisponibleViewModel>> ObtenerEmpresasAsync(CancellationToken cancellationToken)
    {
        var aspNetUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(aspNetUserId))
        {
            return [];
        }

        var empresas = await empresaRepository.ListarPorUsuarioAsync(aspNetUserId, cancellationToken);
        return empresas
            .Select(x => new EmpresaDisponibleViewModel
            {
                IdEmpresa = x.IdEmpresa,
                CodigoEmpresa = x.CodigoEmpresa,
                RazonSocial = x.RazonSocial,
                NombreComercial = x.NombreComercial,
                Ruc = x.Ruc,
                EsEmpresaPredeterminada = x.EsEmpresaPredeterminada
            })
            .ToList();
    }
}
