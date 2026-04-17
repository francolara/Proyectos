using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class PerfilPublicoController(ISportCenterStoredProcedureService spService) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(string? tab = null)
    {
        ViewData["PublicFullWidth"] = true;
        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId)) return Challenge();

        var perfil = await spService.UsuariosPublicosObtenerPerfilAsync(usuarioId) ?? new UsuarioPublicoPerfilViewModel
        {
            UsuarioId = usuarioId,
            TipoDocumento = "0"
        };

        await CargarCombosAsync(perfil);
        var reservas = await spService.UsuariosPublicosReservasListarAsync(usuarioId);

        ViewData["Tab"] = string.IsNullOrWhiteSpace(tab) ? "datos" : tab.Trim().ToLowerInvariant();
        return View(new PerfilPublicoIndexViewModel
        {
            Perfil = perfil,
            Reservas = reservas
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarDatos(PerfilPublicoIndexViewModel vm)
    {
        ViewData["PublicFullWidth"] = true;
        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId)) return Challenge();

        vm.Perfil.UsuarioId = usuarioId;
        vm.Perfil.TipoDocumento = string.IsNullOrWhiteSpace(vm.Perfil.TipoDocumento) ? "0" : vm.Perfil.TipoDocumento.Trim();
        vm.Perfil.NumeroDocumento = string.IsNullOrWhiteSpace(vm.Perfil.NumeroDocumento) ? null : vm.Perfil.NumeroDocumento.Trim();
        vm.Perfil.Nombres = (vm.Perfil.Nombres ?? string.Empty).Trim();
        vm.Perfil.Apellidos = (vm.Perfil.Apellidos ?? string.Empty).Trim();
        vm.Perfil.NombreEquipo = string.IsNullOrWhiteSpace(vm.Perfil.NombreEquipo) ? null : vm.Perfil.NombreEquipo.Trim();
        vm.Perfil.Telefono = string.IsNullOrWhiteSpace(vm.Perfil.Telefono) ? null : vm.Perfil.Telefono.Trim();
        vm.Perfil.Correo = string.IsNullOrWhiteSpace(vm.Perfil.Correo) ? null : vm.Perfil.Correo.Trim();
        vm.Perfil.CodigoUbigeo = string.IsNullOrWhiteSpace(vm.Perfil.CodigoUbigeo) ? null : vm.Perfil.CodigoUbigeo.Trim();

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(vm.Perfil);
            vm.Reservas = await spService.UsuariosPublicosReservasListarAsync(usuarioId);
            ViewData["Tab"] = "datos";
            return View("Index", vm);
        }

        await spService.UsuariosPublicosGuardarPerfilAsync(vm.Perfil, User.Identity?.Name ?? "perfil-publico");
        TempData["PerfilPublicoOk"] = "Datos personales actualizados correctamente.";
        return RedirectToAction(nameof(Index), new { tab = "datos" });
    }

    private async Task CargarCombosAsync(UsuarioPublicoPerfilViewModel perfil)
    {
        perfil.TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        if (perfil.TiposDocumento.All(x => x.Value != "0"))
            perfil.TiposDocumento.Insert(0, new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem("Doc. trib. no dom. sin RUC (0)", "0"));

        perfil.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();
        perfil.ProvinciasUbigeo = !string.IsNullOrWhiteSpace(perfil.CodigoDepartamento) && perfil.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(perfil.CodigoDepartamento)
            : new();
        perfil.DistritosUbigeo = !string.IsNullOrWhiteSpace(perfil.CodigoProvincia) && perfil.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(perfil.CodigoProvincia)
            : new();
    }
}
