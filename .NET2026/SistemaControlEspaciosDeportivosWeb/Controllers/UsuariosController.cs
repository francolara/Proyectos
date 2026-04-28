using System.Security.Cryptography;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Linq;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class UsuariosController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService,
    UserManager<ApplicationUser> userManager)
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
            SedeIdAsignada = baseVm.SedeIdAsignada,
            EsAdministrador = baseVm.EsAdministrador,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            AsignarForm = new UsuarioNegocioAsignarFormViewModel { NegocioId = negocioId }
        };

        var sedeFiltro = AplicarSedeAsignada(baseVm, null);
        vm.Usuarios = await spService.UsuariosNegocioListarAsync(negocioId, sedeFiltro);
        vm.Sedes = await spService.EspaciosComboSedesAsync(negocioId, sedeFiltro);
        vm.Roles = await ObtenerRolesAsync();
        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
        {
            vm.AsignarForm.SedeId = baseVm.SedeIdAsignada;
        }

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
            if (!baseVm.EsAdministrador && model.RolNegocio == 1)
                throw new InvalidOperationException("Solo un administrador puede asignar el rol Administrador.");

            var limites = await spService.NegocioObtenerLimitesOperativosAsync(model.NegocioId);
            var usuariosActivos = (await spService.UsuariosNegocioListarAsync(model.NegocioId))
                .Count(x => x.Activo);
            if (usuariosActivos >= limites.UsuariosPermitidos)
                throw new InvalidOperationException($"Limite de usuarios alcanzado. Tu plan actual permite hasta {limites.UsuariosPermitidos} usuario(s) activos. Para continuar, solicita una ampliacion al administrador de plataforma.");

            var sedeAsignada = baseVm.EsAdministrador ? model.SedeId : baseVm.SedeIdAsignada;
            if (RolRequiereSede(model.RolNegocio) && !sedeAsignada.HasValue)
                throw new InvalidOperationException("Debes seleccionar una sede para usuarios no administradores.");
            if (!SedePermitida(baseVm, sedeAsignada))
                throw new InvalidOperationException("No puedes asignar una sede distinta a la que tienes permitida.");

            var (usuarioSistema, creadoNuevo, claveTemporal) = await ObtenerOCrearUsuarioSistemaAsync(model.Correo, model.NombreUsuario);

            await spService.UsuariosNegocioAsignarPorCorreoAsync(model.NegocioId, model.Correo, model.RolNegocio, sedeAsignada, User.Identity?.Name ?? "sistema");
            TempData["UsuariosMsg"] = creadoNuevo
                ? $"Usuario creado y asignado correctamente. Usuario: {usuarioSistema.UserName}. Clave temporal: {claveTemporal}"
                : "Usuario asignado correctamente.";
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
            var sedeFiltro = AplicarSedeAsignada(baseVm, null);
            var usuarioObjetivo = (await spService.UsuariosNegocioListarAsync(model.NegocioId, sedeFiltro))
                .FirstOrDefault(x => x.UsuarioNegocioId == model.UsuarioNegocioId);
            if (usuarioObjetivo is null)
                throw new InvalidOperationException("No puedes editar usuarios fuera de tu sede asignada.");

            if (!baseVm.EsAdministrador && model.RolNegocio == 1)
                throw new InvalidOperationException("Solo un administrador puede asignar el rol Administrador.");

            var sedeAsignada = baseVm.EsAdministrador ? model.SedeId : baseVm.SedeIdAsignada;
            if (RolRequiereSede(model.RolNegocio) && !sedeAsignada.HasValue)
                throw new InvalidOperationException("Debes seleccionar una sede para usuarios no administradores.");
            if (!SedePermitida(baseVm, sedeAsignada))
                throw new InvalidOperationException("No puedes asignar una sede distinta a la que tienes permitida.");

            await spService.UsuariosNegocioActualizarRolAsync(model.NegocioId, model.UsuarioNegocioId, model.RolNegocio, sedeAsignada, User.Identity?.Name ?? "sistema");
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
            var sedeFiltro = AplicarSedeAsignada(baseVm, null);
            var usuarioObjetivo = (await spService.UsuariosNegocioListarAsync(negocioId, sedeFiltro))
                .FirstOrDefault(x => x.UsuarioNegocioId == usuarioNegocioId);
            if (usuarioObjetivo is null)
                throw new InvalidOperationException("No puedes desactivar usuarios fuera de tu sede asignada.");

            await spService.UsuariosNegocioDesactivarAsync(negocioId, usuarioNegocioId, User.Identity?.Name ?? "sistema");
            TempData["UsuariosMsg"] = "Usuario desactivado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["UsuariosErr"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ResetearClaveTemporal(int negocioId, int usuarioNegocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "USUARIOS");
        if (baseVm is null || !baseVm.PuedeEditar || !baseVm.EsAdministrador)
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Solo un administrador puede resetear claves." });

        try
        {
            var sedeFiltro = AplicarSedeAsignada(baseVm, null);
            var usuarioObjetivo = (await spService.UsuariosNegocioListarAsync(negocioId, sedeFiltro))
                .FirstOrDefault(x => x.UsuarioNegocioId == usuarioNegocioId);
            if (usuarioObjetivo is null)
                throw new InvalidOperationException("No se encontro el usuario a resetear.");

            var usuarioSistema = await userManager.FindByIdAsync(usuarioObjetivo.UsuarioId);
            if (usuarioSistema is null)
                throw new InvalidOperationException("El usuario no existe en Identity.");

            var claveTemporal = GenerarClaveTemporal();
            var token = await userManager.GeneratePasswordResetTokenAsync(usuarioSistema);
            var result = await userManager.ResetPasswordAsync(usuarioSistema, token, claveTemporal);
            if (!result.Succeeded)
            {
                var errores = string.Join(" ", result.Errors.Select(e => e.Description));
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(errores)
                    ? "No se pudo resetear la clave."
                    : errores);
            }

            TempData["UsuariosMsg"] = $"Clave temporal generada para {usuarioSistema.Email}: {claveTemporal}";
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

        var usuarios = await spService.UsuariosNegocioListarAsync(negocioId, AplicarSedeAsignada(baseVm, null));
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
            Roles = await ObtenerRolesAsync(),
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
            var usuarios = await spService.UsuariosNegocioListarAsync(negocioId, AplicarSedeAsignada(baseVm, null));
            if (!usuarios.Any(x => x.UsuarioNegocioId == usuarioNegocioId))
                throw new InvalidOperationException("No puedes modificar permisos fuera de tu sede asignada.");

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

    private async Task<List<SelectListItem>> ObtenerRolesAsync()
    {
        return await spService.UsuariosNegocioRolesListarAsync();
    }

    private static bool RolRequiereSede(int rolNegocio) => rolNegocio != 1;

    private async Task<(ApplicationUser Usuario, bool Creado, string? ClaveTemporal)> ObtenerOCrearUsuarioSistemaAsync(string correo, string? nombreUsuario)
    {
        var correoNormalizado = (correo ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(correoNormalizado))
            throw new InvalidOperationException("El correo es obligatorio.");

        var usuarioExistente = await userManager.FindByEmailAsync(correoNormalizado);
        if (usuarioExistente is not null)
            return (usuarioExistente, false, null);

        var nombreUsuarioLimpio = (nombreUsuario ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(nombreUsuarioLimpio))
            throw new InvalidOperationException("Para crear un usuario nuevo, ingresa nombres y apellidos.");

        var usuarioNuevo = new ApplicationUser
        {
            UserName = correoNormalizado,
            Email = correoNormalizado,
            EmailConfirmed = false,
            Nombres = nombreUsuarioLimpio,
            Apellidos = string.Empty
        };

        var claveTemporal = GenerarClaveTemporal();
        var result = await userManager.CreateAsync(usuarioNuevo, claveTemporal);
        if (!result.Succeeded)
        {
            var errores = string.Join(" ", result.Errors.Select(e => e.Description));
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(errores)
                ? "No se pudo crear el usuario en el sistema."
                : errores);
        }

        return (usuarioNuevo, true, claveTemporal);
    }

    private static string GenerarClaveTemporal()
    {
        const string mayus = "ABCDEFGHJKLMNPQRSTUVWXYZ";
        const string minus = "abcdefghijkmnopqrstuvwxyz";
        const string digitos = "23456789";
        const string simbolos = "!@#$%&*";

        char pick(string chars) => chars[RandomNumberGenerator.GetInt32(chars.Length)];

        var bloque = new char[6];
        for (var i = 0; i < bloque.Length; i++)
            bloque[i] = pick(mayus + minus + digitos);

        return $"Tmp{pick(mayus)}{pick(minus)}{pick(digitos)}{pick(simbolos)}{new string(bloque)}";
    }

}
