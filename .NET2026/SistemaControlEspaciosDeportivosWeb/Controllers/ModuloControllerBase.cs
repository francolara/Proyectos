using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Security.Claims;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public abstract class ModuloControllerBase(IModuloPermisoService moduloPermisoService) : Controller
{
    protected async Task<ModuloBaseViewModel?> ObtenerBaseAsync(int negocioId, string moduloCodigo)
    {
        ViewData["AdminShell"] = true;
        ViewData["AdminNegocioId"] = negocioId;

        var contexto = await moduloPermisoService.ObtenerContextoAsync(User, negocioId, moduloCodigo);
        if (!contexto.Autorizado)
        {
            return new ModuloBaseViewModel
            {
                NegocioId = negocioId,
                ModuloCodigo = moduloCodigo,
                Mensaje = contexto.Mensaje
            };
        }

        return new ModuloBaseViewModel
        {
            NegocioId = contexto.NegocioId,
            NegocioNombre = contexto.NegocioNombre,
            RolActual = contexto.RolActual,
            ModuloCodigo = contexto.ModuloCodigo,
            ModuloNombre = contexto.ModuloNombre,
            PuedeCrear = contexto.PuedeCrear,
            PuedeEditar = contexto.PuedeEditar,
            PuedeEliminar = contexto.PuedeEliminar
        };
    }

    protected IActionResult SinAcceso(ModuloBaseViewModel vm) => View("~/Views/Shared/ModuloSinAcceso.cshtml", vm);

    protected async Task<int?> ResolverNegocioIdAsync(int? negocioId, ISportCenterStoredProcedureService spService)
    {
        if (negocioId.HasValue && negocioId.Value > 0) return negocioId.Value;

        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId)) return null;

        var negocios = await spService.PanelListarNegociosUsuarioAsync(usuarioId);
        return negocios.FirstOrDefault()?.NegocioId;
    }
}
