using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Usuarios;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("USUARIOS")]
public class UsuariosController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    UserManager<IdentityUser> userManager) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        CargarViewDataComun();
        return View(await ConstruirIndexViewModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarUsuario(UsuariosIndexViewModel model, CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        model.Formulario.EmpresasSeleccionadas = model.Formulario.EmpresasSeleccionadas.Distinct().ToList();
        ValidarFormularioUsuario(model.Formulario, ModelState);
        if (!ModelState.IsValid)
        {
            CargarViewDataComun();
            var recargado = await ConstruirIndexViewModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken);
            recargado.Formulario = model.Formulario;
            return View("Index", recargado);
        }

        var email = model.Formulario.Correo.Trim();
        var user = await userManager.FindByEmailAsync(email);
        var usuarioYaVinculado = user is not null
            && (await cuentaAdministradoraRepository.ListarUsuariosCuentaAdministradoraAsync(
                cuenta.Value.idCuentaAdministradora,
                cancellationToken))
            .Any(x => string.Equals(x.AspNetUserId, user.Id, StringComparison.Ordinal));

        if (!usuarioYaVinculado)
        {
            var aspNetUserIdActual = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var contexto = string.IsNullOrWhiteSpace(aspNetUserIdActual)
                ? null
                : await cuentaAdministradoraRepository.ObtenerContextoLoginUsuarioAsync(aspNetUserIdActual, cancellationToken);
            var usuariosActivos = await cuentaAdministradoraRepository.ListarUsuariosCuentaAdministradoraAsync(
                cuenta.Value.idCuentaAdministradora,
                cancellationToken);

            if (contexto?.UsuariosPermitidos is int usuariosPermitidos
                && usuariosActivos.Count >= usuariosPermitidos)
            {
                ModelState.AddModelError(
                    string.Empty,
                    $"La cuenta alcanzo el limite de {usuariosPermitidos} usuario(s) permitido por su suscripcion.");
                CargarViewDataComun();
                var recargado = await ConstruirIndexViewModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken);
                recargado.Formulario = model.Formulario;
                return View("Index", recargado);
            }
        }

        var usuarioCreado = false;
        if (user is null)
        {
            if (string.IsNullOrWhiteSpace(model.Formulario.PasswordTemporal))
            {
                ModelState.AddModelError(nameof(model.Formulario.PasswordTemporal), "Ingrese una contrasena temporal para crear el usuario.");
                CargarViewDataComun();
                var recargado = await ConstruirIndexViewModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken);
                recargado.Formulario = model.Formulario;
                return View("Index", recargado);
            }

            user = new IdentityUser
            {
                UserName = email,
                Email = email,
                EmailConfirmed = true
            };

            var createResult = await userManager.CreateAsync(user, model.Formulario.PasswordTemporal!.Trim());
            if (!createResult.Succeeded)
            {
                foreach (var error in createResult.Errors)
                {
                    ModelState.AddModelError(string.Empty, IdentityErrorTranslator.Translate(error));
                }

                CargarViewDataComun();
                var recargado = await ConstruirIndexViewModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken);
                recargado.Formulario = model.Formulario;
                return View("Index", recargado);
            }

            usuarioCreado = true;
            var temporaryPasswordClaimResult = await userManager.AddClaimAsync(
                user,
                new Claim(
                    TemporaryPasswordFlowConstants.RequirePasswordChangeClaimType,
                    TemporaryPasswordFlowConstants.RequirePasswordChangeClaimValue));

            if (!temporaryPasswordClaimResult.Succeeded)
            {
                await userManager.DeleteAsync(user);

                foreach (var error in temporaryPasswordClaimResult.Errors)
                {
                    ModelState.AddModelError(string.Empty, IdentityErrorTranslator.Translate(error));
                }

                CargarViewDataComun();
                var recargado = await ConstruirIndexViewModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken);
                recargado.Formulario = model.Formulario;
                return View("Index", recargado);
            }
        }

        try
        {
            await cuentaAdministradoraRepository.AsignarUsuarioCuentaAdministradoraAsync(new AsignarUsuarioCuentaAdministradoraRequest
            {
                AspNetUserId = user.Id,
                IdCuentaAdministradora = cuenta.Value.idCuentaAdministradora,
                RolCuenta = model.Formulario.RolCuenta,
                EsCuentaPredeterminada = false,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);
        }
        catch (SqlException ex) when (ex.Number == 50000)
        {
            if (usuarioCreado)
            {
                await userManager.DeleteAsync(user);
            }

            ModelState.AddModelError(string.Empty, ex.Message);
            CargarViewDataComun();
            var recargado = await ConstruirIndexViewModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken);
            recargado.Formulario = model.Formulario;
            return View("Index", recargado);
        }

        await cuentaAdministradoraRepository.GuardarPerfilUsuarioAsync(new UsuarioPerfilRequest
        {
            AspNetUserId = user.Id,
            NombreCompleto = model.Formulario.NombreCompleto.Trim(),
            Telefono = LimpiarTelefono(model.Formulario.Telefono),
            CorreoReferencia = email,
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        await SincronizarEmpresasUsuarioAsync(
            cuenta.Value.idCuentaAdministradora,
            user.Id,
            model.Formulario.EmpresasSeleccionadas,
            null,
            cancellationToken);

        TempData["SuccessMessage"] = "El usuario fue vinculado a la cuenta administradora.";
        return RedirectToAction(nameof(Permisos), new { aspNetUserId = user.Id });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModulePermission("USUARIOS", ModulePermissionOperation.Delete)]
    public async Task<IActionResult> Desactivar(string aspNetUserId, CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await cuentaAdministradoraRepository.DesactivarUsuarioCuentaAdministradoraAsync(
            aspNetUserId,
            cuenta.Value.idCuentaAdministradora,
            User.Identity?.Name,
            cancellationToken);

        TempData["SuccessMessage"] = "El acceso del usuario fue desactivado.";
        return RedirectToAction(nameof(Index));
    }

    [HttpGet]
    public async Task<IActionResult> Permisos(string aspNetUserId, int? idEmpresa, CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var model = await ConstruirPermisosViewModelAsync(cuenta.Value.idCuentaAdministradora, aspNetUserId, idEmpresa, cancellationToken);
        if (model is null)
        {
            TempData["ErrorMessage"] = "No se encontro el usuario solicitado dentro de la cuenta administradora.";
            return RedirectToAction(nameof(Index));
        }

        CargarViewDataComun();
        ViewData["Title"] = "Permisos de usuario";
        ViewBag.RoleOptions = CrearRoleOptions(model.RolCuenta);
        ViewBag.EmpresaFiltroOptions = model.EmpresasDisponibles
            .Where(x => x.IdUsuarioEmpresa.HasValue)
            .Select(x => new SelectListItem($"{x.CodigoEmpresa} - {x.RazonSocial}", x.IdEmpresa.ToString(), x.IdEmpresa == model.IdEmpresaSeleccionada))
            .ToList();
        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ActualizarAcceso(UsuarioPermisosViewModel model, CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        model.EmpresasSeleccionadas = model.EmpresasSeleccionadas.Distinct().ToList();
        if (string.IsNullOrWhiteSpace(model.RolCuenta))
        {
            ModelState.AddModelError(nameof(model.RolCuenta), "Seleccione el rol del usuario.");
        }
        if (!ModelState.IsValid)
        {
            return RedirectToAction(nameof(Permisos), new { aspNetUserId = model.AspNetUserId, idEmpresa = model.IdEmpresaSeleccionada });
        }

        var accesoActual = (await cuentaAdministradoraRepository
            .ListarUsuariosCuentaAdministradoraAsync(cuenta.Value.idCuentaAdministradora, cancellationToken))
            .FirstOrDefault(x => string.Equals(x.AspNetUserId, model.AspNetUserId, StringComparison.Ordinal));

        if (accesoActual is null)
        {
            TempData["ErrorMessage"] = "No se encontro el acceso del usuario dentro de la cuenta administradora.";
            return RedirectToAction(nameof(Index));
        }

        await cuentaAdministradoraRepository.AsignarUsuarioCuentaAdministradoraAsync(new AsignarUsuarioCuentaAdministradoraRequest
        {
            AspNetUserId = model.AspNetUserId,
            IdCuentaAdministradora = cuenta.Value.idCuentaAdministradora,
            RolCuenta = model.RolCuenta,
            EsCuentaPredeterminada = accesoActual.EsCuentaPredeterminada,
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        await SincronizarEmpresasUsuarioAsync(
            cuenta.Value.idCuentaAdministradora,
            model.AspNetUserId,
            model.EmpresasSeleccionadas,
            null,
            cancellationToken);

        TempData["SuccessMessage"] = "Se actualizaron el rol y las empresas del usuario.";
        return RedirectToAction(nameof(Permisos), new { aspNetUserId = model.AspNetUserId, idEmpresa = model.IdEmpresaSeleccionada });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarPermisoCuenta(string aspNetUserId, int idUsuarioCuentaAdministradora, int idModuloSistema, string puedeVer, string puedeCrear, string puedeEditar, string puedeEliminar, CancellationToken cancellationToken)
    {
        await cuentaAdministradoraRepository.GuardarUsuarioCuentaPermisoAsync(new GuardarUsuarioCuentaPermisoRequest
        {
            IdUsuarioCuentaAdministradora = idUsuarioCuentaAdministradora,
            IdModuloSistema = idModuloSistema,
            PuedeVer = ResolverOverrideBool(puedeVer),
            PuedeCrear = ResolverOverrideBool(puedeCrear),
            PuedeEditar = ResolverOverrideBool(puedeEditar),
            PuedeEliminar = ResolverOverrideBool(puedeEliminar),
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["SuccessMessage"] = "Permiso de cuenta actualizado.";
        return RedirectToAction(nameof(Permisos), new { aspNetUserId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarPermisoEmpresa(string aspNetUserId, int idUsuarioEmpresa, int idEmpresa, int idModuloSistema, string puedeVer, string puedeCrear, string puedeEditar, string puedeEliminar, CancellationToken cancellationToken)
    {
        await cuentaAdministradoraRepository.GuardarUsuarioEmpresaPermisoAsync(new GuardarUsuarioEmpresaPermisoRequest
        {
            IdUsuarioEmpresa = idUsuarioEmpresa,
            IdModuloSistema = idModuloSistema,
            PuedeVer = ResolverOverrideBool(puedeVer),
            PuedeCrear = ResolverOverrideBool(puedeCrear),
            PuedeEditar = ResolverOverrideBool(puedeEditar),
            PuedeEliminar = ResolverOverrideBool(puedeEliminar),
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["SuccessMessage"] = "Permiso por empresa actualizado.";
        return RedirectToAction(nameof(Permisos), new { aspNetUserId, idEmpresa });
    }

    private async Task<UsuariosIndexViewModel> ConstruirIndexViewModelAsync(int idCuentaAdministradora, CancellationToken cancellationToken)
    {
        var configuracion = await cuentaAdministradoraRepository.ObtenerConfiguracionCuentaAdministradoraAsync(idCuentaAdministradora, cancellationToken);
        var usuarios = await cuentaAdministradoraRepository.ListarUsuariosCuentaAdministradoraAsync(idCuentaAdministradora, cancellationToken);
        var empresas = await cuentaAdministradoraRepository.ListarEmpresasCuentaAdministradoraAsync(idCuentaAdministradora, cancellationToken);

        return new UsuariosIndexViewModel
        {
            IdCuentaAdministradora = idCuentaAdministradora,
            CodigoCuenta = configuracion?.CodigoCuenta ?? string.Empty,
            NombreCuenta = configuracion?.NombreCuenta ?? string.Empty,
            Usuarios = usuarios.Select(x => new UsuarioCuentaListaItemViewModel
            {
                IdUsuarioCuentaAdministradora = x.IdUsuarioCuentaAdministradora,
                AspNetUserId = x.AspNetUserId,
                CorreoUsuario = x.CorreoUsuario,
                NombreCompleto = string.IsNullOrWhiteSpace(x.NombreCompleto) ? x.CorreoUsuario : x.NombreCompleto,
                Telefono = x.Telefono,
                RolCuenta = x.RolCuenta,
                EsCuentaPredeterminada = x.EsCuentaPredeterminada,
                CantidadEmpresasAsignadas = x.CantidadEmpresasAsignadas,
                EmpresasAsignadas = x.EmpresasAsignadas ?? "Sin empresas asignadas"
            }).ToList(),
            EmpresasDisponibles = empresas.Select(x => new UsuarioEmpresaAsignacionViewModel
            {
                IdEmpresa = x.IdEmpresa,
                CodigoEmpresa = x.CodigoEmpresa,
                RazonSocial = x.RazonSocial
            }).ToList()
        };
    }

    private async Task<UsuarioPermisosViewModel?> ConstruirPermisosViewModelAsync(int idCuentaAdministradora, string aspNetUserId, int? idEmpresa, CancellationToken cancellationToken)
    {
        var usuarios = await cuentaAdministradoraRepository.ListarUsuariosCuentaAdministradoraAsync(idCuentaAdministradora, cancellationToken);
        var usuario = usuarios.FirstOrDefault(x => string.Equals(x.AspNetUserId, aspNetUserId, StringComparison.Ordinal));
        if (usuario is null)
        {
            return null;
        }

        var empresas = (await cuentaAdministradoraRepository.ListarEmpresasUsuarioCuentaAdministradoraAsync(idCuentaAdministradora, aspNetUserId, cancellationToken)).ToList();
        var empresaSeleccionada = idEmpresa.HasValue
            ? empresas.FirstOrDefault(x => x.IdEmpresa == idEmpresa.Value && x.IdUsuarioEmpresa.HasValue)
            : empresas.FirstOrDefault(x => x.EsEmpresaPredeterminada && x.IdUsuarioEmpresa.HasValue)
              ?? empresas.FirstOrDefault(x => x.IdUsuarioEmpresa.HasValue);

        var permisosCuenta = await cuentaAdministradoraRepository.ListarPermisosUsuarioCuentaAsync(usuario.IdUsuarioCuentaAdministradora, cancellationToken);
        var permisosEmpresa = empresaSeleccionada?.IdUsuarioEmpresa is int idUsuarioEmpresa
            ? await cuentaAdministradoraRepository.ListarPermisosUsuarioEmpresaAsync(idUsuarioEmpresa, cancellationToken)
            : [];

        return new UsuarioPermisosViewModel
        {
            IdCuentaAdministradora = idCuentaAdministradora,
            IdUsuarioCuentaAdministradora = usuario.IdUsuarioCuentaAdministradora,
            AspNetUserId = usuario.AspNetUserId,
            CorreoUsuario = usuario.CorreoUsuario,
            NombreCompleto = string.IsNullOrWhiteSpace(usuario.NombreCompleto) ? usuario.CorreoUsuario : usuario.NombreCompleto,
            RolCuenta = usuario.RolCuenta,
            EsCuentaPredeterminada = usuario.EsCuentaPredeterminada,
            IdEmpresaPredeterminada = empresas.FirstOrDefault(x => x.EsEmpresaPredeterminada)?.IdEmpresa,
            EmpresasSeleccionadas = empresas.Where(x => x.Asignado).Select(x => x.IdEmpresa).ToList(),
            EmpresasDisponibles = empresas.Select(x => new UsuarioEmpresaAsignacionViewModel
            {
                IdEmpresa = x.IdEmpresa,
                CodigoEmpresa = x.CodigoEmpresa,
                RazonSocial = x.RazonSocial,
                Asignado = x.Asignado,
                EsEmpresaPredeterminada = x.EsEmpresaPredeterminada,
                IdUsuarioEmpresa = x.IdUsuarioEmpresa
            }).ToList(),
            IdEmpresaSeleccionada = empresaSeleccionada?.IdEmpresa,
            IdUsuarioEmpresaSeleccionada = empresaSeleccionada?.IdUsuarioEmpresa,
            NombreEmpresaSeleccionada = empresaSeleccionada?.RazonSocial,
            PermisosCuenta = permisosCuenta.Select(MapearPermiso).ToList(),
            PermisosEmpresa = permisosEmpresa.Select(MapearPermiso).ToList()
        };
    }

    private async Task SincronizarEmpresasUsuarioAsync(int idCuentaAdministradora, string aspNetUserId, IReadOnlyCollection<int> empresasSeleccionadas, int? idEmpresaPredeterminada, CancellationToken cancellationToken)
    {
        var asignadas = (await cuentaAdministradoraRepository.ListarEmpresasUsuarioCuentaAdministradoraAsync(idCuentaAdministradora, aspNetUserId, cancellationToken)).ToList();
        var seleccionadas = empresasSeleccionadas.Distinct().ToHashSet();

        foreach (var empresaId in seleccionadas)
        {
            await cuentaAdministradoraRepository.AsignarUsuarioEmpresaAsync(new AsignarUsuarioEmpresaRequest
            {
                AspNetUserId = aspNetUserId,
                IdEmpresa = empresaId,
                EsEmpresaPredeterminada = idEmpresaPredeterminada.HasValue && idEmpresaPredeterminada.Value == empresaId,
                UsuarioRegistro = User.Identity?.Name ?? "sistema"
            }, cancellationToken);
        }

        foreach (var empresa in asignadas.Where(x => x.Asignado && !seleccionadas.Contains(x.IdEmpresa)))
        {
            await cuentaAdministradoraRepository.DesactivarUsuarioEmpresaAsync(
                aspNetUserId,
                empresa.IdEmpresa,
                User.Identity?.Name,
                cancellationToken);
        }
    }

    private async Task<(int idCuentaAdministradora, string? nombreCuenta)?> ResolverCuentaAsync(CancellationToken cancellationToken)
    {
        if (currentCompanyAccessor.TieneEmpresaActiva && currentCompanyAccessor.EmpresaId.HasValue)
        {
            var contextoEmpresa = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(
                currentCompanyAccessor.EmpresaId.Value,
                cancellationToken);

            if (contextoEmpresa is not null)
            {
                return (contextoEmpresa.IdCuentaAdministradora, contextoEmpresa.NombreCuenta);
            }
        }

        var aspNetUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(aspNetUserId))
        {
            return null;
        }

        var contextoLogin = await cuentaAdministradoraRepository.ObtenerContextoLoginUsuarioAsync(aspNetUserId, cancellationToken);
        if (contextoLogin is null || !contextoLogin.TieneAcceso || !contextoLogin.IdCuentaAdministradora.HasValue)
        {
            return null;
        }

        return (contextoLogin.IdCuentaAdministradora.Value, contextoLogin.NombreCuenta);
    }

    private void CargarViewDataComun()
    {
        ViewData["Title"] = "Usuarios";
        ViewData["AdminShell"] = true;
        ViewBag.RoleOptions = CrearRoleOptions();
    }

    private static IReadOnlyCollection<SelectListItem> CrearRoleOptions(string? selectedValue = null)
    {
        return
        [
            new("Administrador de cuenta", "ADMINISTRADORCUENTA", string.Equals(selectedValue, "ADMINISTRADORCUENTA", StringComparison.OrdinalIgnoreCase)),
            new("Supervisor", "SUPERVISOR", string.Equals(selectedValue, "SUPERVISOR", StringComparison.OrdinalIgnoreCase)),
            new("Operador", "OPERADOR", string.Equals(selectedValue, "OPERADOR", StringComparison.OrdinalIgnoreCase)),
            new("Consulta", "CONSULTA", string.Equals(selectedValue, "CONSULTA", StringComparison.OrdinalIgnoreCase))
        ];
    }

    private static void ValidarFormularioUsuario(UsuarioCuentaFormViewModel formulario, ModelStateDictionary modelState)
    {
        if (!formulario.EmpresasSeleccionadas.Any())
        {
            modelState.AddModelError(nameof(formulario.EmpresasSeleccionadas), "Seleccione al menos una empresa para el usuario.");
        }
    }

    private static ModuloPermisoViewModel MapearPermiso(UsuarioCuentaPermisoDto dto)
    {
        return new ModuloPermisoViewModel
        {
            IdModuloSistema = dto.IdModuloSistema,
            CodigoModulo = dto.CodigoModulo,
            NombreModulo = dto.NombreModulo,
            GrupoMenu = dto.GrupoMenu ?? string.Empty,
            PuedeVerOverride = dto.PuedeVerOverride,
            PuedeCrearOverride = dto.PuedeCrearOverride,
            PuedeEditarOverride = dto.PuedeEditarOverride,
            PuedeEliminarOverride = dto.PuedeEliminarOverride,
            PuedeVerEfectivo = dto.PuedeVerEfectivo,
            PuedeCrearEfectivo = dto.PuedeCrearEfectivo,
            PuedeEditarEfectivo = dto.PuedeEditarEfectivo,
            PuedeEliminarEfectivo = dto.PuedeEliminarEfectivo
        };
    }

    private static bool? ResolverOverrideBool(string? value)
    {
        var normalized = (value ?? string.Empty).Trim().ToUpperInvariant();
        return normalized switch
        {
            "1" or "TRUE" or "SI" => true,
            "0" or "FALSE" or "NO" => false,
            _ => null
        };
    }

    private static string? LimpiarTelefono(string? telefono)
    {
        if (string.IsNullOrWhiteSpace(telefono))
        {
            return null;
        }

        return new string(telefono.Where(x => char.IsDigit(x) || x == '+').ToArray());
    }
}
