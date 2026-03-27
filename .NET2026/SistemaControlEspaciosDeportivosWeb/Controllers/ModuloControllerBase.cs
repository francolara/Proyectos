using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public abstract class ModuloControllerBase(IModuloPermisoService moduloPermisoService) : Controller
{
    protected async Task<ModuloBaseViewModel?> ObtenerBaseAsync(int negocioId, string moduloCodigo)
    {
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
}

