using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Empresas;
using System.Security.Claims;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class EmpresaContextoController(
    IEmpresaRepository empresaRepository,
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    UserManager<IdentityUser> userManager,
    IModulePermissionService modulePermissionService) : Controller
{
    [HttpGet]
    [ModulePermission("EMPRESAS", ModulePermissionOperation.View)]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        ViewData["Title"] = "Seleccion de empresa";
        ViewData["PuedeCrearEmpresa"] = await modulePermissionService.CanAccessModuleAsync(
            User,
            "EMPRESAS",
            ModulePermissionOperation.Create,
            cancellationToken);
        ViewData["PuedeEditarEmpresa"] = await modulePermissionService.CanAccessModuleAsync(
            User,
            "EMPRESAS",
            ModulePermissionOperation.Edit,
            cancellationToken);
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
    [ModulePermission("EMPRESAS", ModulePermissionOperation.Create)]
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
    [ModulePermission("EMPRESAS", ModulePermissionOperation.Create)]
    public async Task<IActionResult> RegistrarEmpresaInicial(RegistroEmpresaInicialViewModel model, CancellationToken cancellationToken)
    {
        var empresas = await ObtenerEmpresasAsync(cancellationToken);
        if (!ModelState.IsValid)
        {
            model.EsEmpresaInicial = empresas.Count == 0;
            return View("RegistrarEmpresaInicial", model);
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
            var result = await RegistrarCuentaInicialConReintentoAsync(
                aspNetUserId,
                model.NombreContacto.Trim(),
                telefono,
                email,
                codigoEmpresa,
                razonSocial,
                nombreComercial,
                model.Ruc.Trim(),
                cancellationToken);

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
                return View("RegistrarEmpresaInicial", model);
            }

            var empresasCuenta = await cuentaAdministradoraRepository.ListarEmpresasCuentaAdministradoraAsync(
                contexto.IdCuentaAdministradora,
                cancellationToken);
            var empresasActivas = empresasCuenta.Count(x => x.Estado);
            if (contexto.EmpresasPermitidas.HasValue
                && empresasActivas >= contexto.EmpresasPermitidas.Value)
            {
                ModelState.AddModelError(
                    string.Empty,
                    $"La cuenta alcanzo el limite de {contexto.EmpresasPermitidas.Value} empresa(s) permitido por su suscripcion.");
                model.EsEmpresaInicial = false;
                return View("RegistrarEmpresaInicial", model);
            }

            try
            {
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
            catch (SqlException ex) when (ex.Number == 50000)
            {
                ModelState.AddModelError(string.Empty, ex.Message);
                model.EsEmpresaInicial = false;
                return View("RegistrarEmpresaInicial", model);
            }
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

    [HttpGet]
    [ModulePermission("EMPRESAS", ModulePermissionOperation.Edit)]
    public async Task<IActionResult> Editar(int idEmpresa, CancellationToken cancellationToken)
    {
        var empresa = (await ObtenerEmpresasAsync(cancellationToken))
            .FirstOrDefault(x => x.IdEmpresa == idEmpresa);

        if (empresa is null)
        {
            return NotFound();
        }

        return View(new EditarEmpresaViewModel
        {
            IdEmpresa = empresa.IdEmpresa,
            RazonSocial = empresa.RazonSocial,
            NombreComercial = empresa.NombreComercial,
            Ruc = empresa.Ruc
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModulePermission("EMPRESAS", ModulePermissionOperation.Edit)]
    public async Task<IActionResult> Editar(EditarEmpresaViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid)
        {
            return View(model);
        }

        var aspNetUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(aspNetUserId))
        {
            return Challenge();
        }

        var razonSocial = model.RazonSocial.Trim();
        var nombreComercial = string.IsNullOrWhiteSpace(model.NombreComercial)
            ? razonSocial
            : model.NombreComercial.Trim();
        var ruc = model.Ruc.Trim();

        try
        {
            await empresaRepository.ActualizarAsync(new ActualizarEmpresaRequest
            {
                IdEmpresa = model.IdEmpresa,
                AspNetUserId = aspNetUserId,
                RazonSocial = razonSocial,
                NombreComercial = nombreComercial,
                Ruc = ruc,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);
        }
        catch (SqlException ex) when (ex.Number == 50000)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }

        if (currentCompanyAccessor.EmpresaId == model.IdEmpresa)
        {
            currentCompanyAccessor.EstablecerEmpresa(model.IdEmpresa, razonSocial);
        }

        TempData["SuccessMessage"] = "Los datos de la empresa fueron actualizados correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModulePermission("EMPRESAS", ModulePermissionOperation.View)]
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

    private async Task<RegistroCuentaAdministradoraConEmpresaResult> RegistrarCuentaInicialConReintentoAsync(
        string aspNetUserId,
        string nombreContacto,
        string? telefono,
        string email,
        string codigoEmpresa,
        string razonSocial,
        string nombreComercial,
        string ruc,
        CancellationToken cancellationToken)
    {
        const int maxIntentos = 5;
        SqlException? ultimaExcepcion = null;

        for (var intento = 0; intento < maxIntentos; intento++)
        {
            try
            {
                return await cuentaAdministradoraRepository.RegistrarCuentaConEmpresaAsync(new RegistroCuentaAdministradoraConEmpresaRequest
                {
                    AspNetUserId = aspNetUserId,
                    NombreCompleto = nombreContacto,
                    Telefono = telefono,
                    CorreoReferencia = email,
                    CodigoCuenta = GenerarCodigoCuenta(nombreContacto, email, intento),
                    NombreCuenta = nombreContacto,
                    CodigoEmpresa = codigoEmpresa,
                    RazonSocial = razonSocial,
                    NombreComercial = nombreComercial,
                    Ruc = ruc,
                    DiasPrueba = 30,
                    UsuarioRegistro = email
                }, cancellationToken);
            }
            catch (SqlException ex) when (EsCodigoCuentaDuplicado(ex) && intento < maxIntentos - 1)
            {
                ultimaExcepcion = ex;
            }
        }

        if (ultimaExcepcion is not null)
        {
            throw ultimaExcepcion;
        }

        throw new InvalidOperationException("No se pudo registrar la cuenta administradora inicial.");
    }

    private static bool EsCodigoCuentaDuplicado(SqlException ex)
        => ex.Message.Contains("codigo de cuenta ya existe", StringComparison.OrdinalIgnoreCase);

    private static string GenerarCodigoCuenta(string nombreCuenta, string correo, int intento)
    {
        var baseCodigo = new string(nombreCuenta.Where(char.IsLetterOrDigit).Take(10).ToArray()).ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = new string(correo.Where(char.IsLetterOrDigit).Take(10).ToArray()).ToUpperInvariant();
        }

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = "CTA";
        }

        var sufijo = intento <= 0
            ? DateTime.UtcNow.ToString("ddHHmm")
            : $"{DateTime.UtcNow:HHmm}{intento:0}";
        var codigo = $"{baseCodigo}{sufijo}".ToUpperInvariant();
        return codigo.Length > 20 ? codigo[..20] : codigo;
    }
}
