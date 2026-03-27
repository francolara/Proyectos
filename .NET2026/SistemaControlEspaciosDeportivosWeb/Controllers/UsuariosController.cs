using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Linq;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class UsuariosController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "USUARIOS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new UsuariosNegocioIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Usuarios = await spService.UsuariosNegocioListarAsync(negocioId),
            AsignarForm = new UsuarioNegocioAsignarFormViewModel { NegocioId = negocioId }
        };

        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Asignar(UsuarioNegocioAsignarFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "USUARIOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.UsuariosNegocioAsignarPorCorreoAsync(model.NegocioId, model.Correo, model.RolNegocio, User.Identity?.Name ?? "sistema");
            TempData["UsuariosMsg"] = "Usuario asignado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["UsuariosErr"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CambiarRol(UsuarioNegocioRolFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "USUARIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.UsuariosNegocioActualizarRolAsync(model.NegocioId, model.UsuarioNegocioId, model.RolNegocio, User.Identity?.Name ?? "sistema");
            TempData["UsuariosMsg"] = "Rol actualizado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["UsuariosErr"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Desactivar(int negocioId, int usuarioNegocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "USUARIOS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.UsuariosNegocioDesactivarAsync(negocioId, usuarioNegocioId, User.Identity?.Name ?? "sistema");
            TempData["UsuariosMsg"] = "Usuario desactivado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["UsuariosErr"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    public async Task<IActionResult> Permisos(int negocioId, int usuarioNegocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "USUARIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var usuarios = await spService.UsuariosNegocioListarAsync(negocioId);
        var usuario = usuarios.FirstOrDefault(x => x.UsuarioNegocioId == usuarioNegocioId);
        if (usuario is null) return NotFound();

        var vm = new UsuarioNegocioPermisosViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            UsuarioNegocioId = usuarioNegocioId,
            UsuarioNombre = $"{usuario.Nombres} {usuario.Apellidos}".Trim(),
            UsuarioCorreo = usuario.Correo,
            RolUsuarioNegocio = usuario.RolNegocio,
            Roles = ObtenerRoles(),
            Modulos = await spService.UsuariosNegocioPermisosListarAsync(negocioId, usuarioNegocioId)
        };

        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarPermisos(int negocioId, int usuarioNegocioId, List<UsuarioNegocioPermisoModuloViewModel> modulos)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "USUARIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            foreach (var modulo in modulos)
            {
                await spService.UsuariosNegocioPermisoGuardarAsync(negocioId, usuarioNegocioId, modulo, User.Identity?.Name ?? "sistema");
            }
            TempData["UsuariosMsg"] = "Permisos guardados correctamente.";
        }
        catch (Exception ex)
        {
            TempData["UsuariosErr"] = ex.Message;
        }

        return RedirectToAction(nameof(Permisos), new { negocioId, usuarioNegocioId });
    }

    private static List<SelectListItem> ObtenerRoles()
    {
        return new()
        {
            new("Administrador", "1"),
            new("Trabajador", "2"),
            new("Recepcion", "3"),
            new("Caja", "4"),
            new("Supervisor", "5")
        };
    }
}
