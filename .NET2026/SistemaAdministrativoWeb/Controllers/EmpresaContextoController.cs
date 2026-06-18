using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Empresas;
using System.Security.Claims;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class EmpresaContextoController(
    IEmpresaRepository empresaRepository,
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    UserManager<IdentityUser> userManager) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        ViewData["Title"] = "Seleccion de empresa";
        var model = new SeleccionEmpresaViewModel
        {
            Empresas = await ObtenerEmpresasAsync(cancellationToken)
        };

        if (model.Empresas.Count == 0)
        {
            return RedirectToAction(nameof(RegistrarEmpresaInicial));
        }

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

    [HttpGet]
    public async Task<IActionResult> RegistrarEmpresaInicial()
    {
        var empresas = await ObtenerEmpresasAsync(HttpContext.RequestAborted);
        var user = await userManager.GetUserAsync(User);
        var model = new RegistroEmpresaInicialViewModel
        {
            Correo = user?.Email ?? User.Identity?.Name ?? string.Empty,
            EsEmpresaInicial = empresas.Count == 0
        };

        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> RegistrarEmpresaInicial(RegistroEmpresaInicialViewModel model, CancellationToken cancellationToken)
    {
        var empresas = await ObtenerEmpresasAsync(cancellationToken);
        if (!ModelState.IsValid)
        {
            model.EsEmpresaInicial = empresas.Count == 0;
            return View(model);
        }

        var aspNetUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(aspNetUserId))
        {
            return Challenge();
        }

        var user = await userManager.GetUserAsync(User);
        var email = (user?.Email ?? model.Correo).Trim();
        var razonSocial = model.RazonSocial.Trim();
        var nombreComercial = string.IsNullOrWhiteSpace(model.NombreComercial) ? razonSocial : model.NombreComercial.Trim();
        var codigoEmpresa = GenerarCodigoEmpresa(razonSocial, model.Ruc);
        var telefono = LimpiarTelefono(model.Telefono);

        int idEmpresa;
        if (empresas.Count == 0)
        {
            var result = await cuentaAdministradoraRepository.RegistrarCuentaConEmpresaAsync(new RegistroCuentaAdministradoraConEmpresaRequest
            {
                AspNetUserId = aspNetUserId,
                NombreCompleto = model.NombreContacto.Trim(),
                Telefono = telefono,
                CorreoReferencia = email,
                CodigoCuenta = GenerarCodigoCuenta(model.NombreContacto, email),
                NombreCuenta = model.NombreContacto.Trim(),
                CodigoEmpresa = codigoEmpresa,
                RazonSocial = razonSocial,
                NombreComercial = nombreComercial,
                Ruc = model.Ruc.Trim(),
                DiasPrueba = 30,
                UsuarioRegistro = email
            }, cancellationToken);

            idEmpresa = result.IdEmpresa;
        }
        else
        {
            var empresaBaseId = currentCompanyAccessor.EmpresaId ?? empresas.FirstOrDefault(x => x.EsEmpresaPredeterminada)?.IdEmpresa ?? empresas[0].IdEmpresa;
            var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(empresaBaseId, cancellationToken);
            if (contexto is null)
            {
                ModelState.AddModelError(string.Empty, "No se pudo resolver la cuenta administradora para registrar la nueva empresa.");
                model.EsEmpresaInicial = false;
                return View(model);
            }

            var result = await cuentaAdministradoraRepository.RegistrarEmpresaCuentaAsync(new RegistroEmpresaCuentaAdministradoraRequest
            {
                IdCuentaAdministradora = contexto.IdCuentaAdministradora,
                AspNetUserId = aspNetUserId,
                CodigoEmpresa = codigoEmpresa,
                RazonSocial = razonSocial,
                NombreComercial = nombreComercial,
                Ruc = model.Ruc.Trim(),
                EsEmpresaPredeterminada = false,
                UsuarioRegistro = email
            }, cancellationToken);

            idEmpresa = result.IdEmpresa;
        }

        if (user is not null && !await userManager.IsInRoleAsync(user, "AdministradorEmpresa"))
        {
            await userManager.AddToRoleAsync(user, "AdministradorEmpresa");
        }

        currentCompanyAccessor.EstablecerEmpresa(idEmpresa, razonSocial);
        TempData["SuccessMessage"] = empresas.Count == 0
            ? "La empresa inicial fue registrada correctamente."
            : "La nueva empresa fue registrada correctamente.";
        return RedirectToAction(nameof(Index));
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

    private static string? LimpiarTelefono(string? telefono)
    {
        if (string.IsNullOrWhiteSpace(telefono))
        {
            return null;
        }

        return new string(telefono.Where(x => char.IsDigit(x) || x == '+').ToArray());
    }

    private static string GenerarCodigoEmpresa(string razonSocial, string ruc)
    {
        var baseCodigo = string.IsNullOrWhiteSpace(ruc)
            ? new string(razonSocial.Where(char.IsLetterOrDigit).Take(8).ToArray()).ToUpperInvariant()
            : ruc.Trim();

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = $"EMP{DateTime.UtcNow:HHmmss}";
        }

        return baseCodigo.Length > 20 ? baseCodigo[..20] : baseCodigo;
    }

    private static string GenerarCodigoCuenta(string nombreCuenta, string correo)
    {
        var baseCodigo = new string(nombreCuenta.Where(char.IsLetterOrDigit).Take(12).ToArray()).ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = new string(correo.Where(char.IsLetterOrDigit).Take(12).ToArray()).ToUpperInvariant();
        }

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = $"CTA{DateTime.UtcNow:HHmmss}";
        }

        return baseCodigo.Length > 20 ? baseCodigo[..20] : baseCodigo;
    }
}
