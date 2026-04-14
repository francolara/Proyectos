using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Diagnostics;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class HomeController(
    ISportCenterStoredProcedureService spService,
    INotificacionEmailService notificacionEmailService,
    UserManager<ApplicationUser> userManager,
    SignInManager<ApplicationUser> signInManager) : Controller
{
    private const string CaptchaSoftwareClubesSessionKey = "CaptchaSoftwareClubes";
    public async Task<IActionResult> Index(
        DateOnly? fecha,
        TimeOnly? horaInicio,
        TimeOnly? horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId)
    {
        var vm = await ConstruirHomeVmAsync(fecha, horaInicio, horaFin, codigoDepartamento, codigoProvincia, codigoUbigeo, tipoDeporteId);
        vm.MensajeSolicitud = TempData["MensajeSolicitud"]?.ToString();
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SolicitarReservaPublica(SolicitudReservaPublicaFormViewModel model)
    {
        if (model.HoraFin <= model.HoraInicio)
            ModelState.AddModelError(string.Empty, "La hora fin debe ser mayor que la hora inicio.");

        if (!ModelState.IsValid)
        {
            var vmError = await ConstruirHomeVmAsync(model.Fecha, model.HoraInicio, model.HoraFin, model.CodigoDepartamento, model.CodigoProvincia, model.CodigoUbigeo, model.TipoDeporteId);
            return View("Index", vmError);
        }

        try
        {
            var codigo = await spService.HomeSolicitarReservaPublicaAsync(model);
            var payloadEmail = await spService.HomeObtenerSolicitudParaNotificacionAsync(codigo);
            if (payloadEmail is not null)
            {
                var enviado = false;
                try
                {
                    enviado = await notificacionEmailService.EnviarSolicitudRecibidaAsync(payloadEmail);
                }
                catch
                {
                    enviado = false;
                }

                if (enviado)
                {
                    await spService.HomeMarcarSolicitudNotificadaAsync(codigo);
                }
            }

            TempData["MensajeSolicitud"] = $"Solicitud registrada correctamente. Codigo: {codigo}. Te contactaremos para confirmacion.";
            return RedirectToAction(nameof(Index), new
            {
                fecha = model.Fecha,
                horaInicio = model.HoraInicio,
                horaFin = model.HoraFin,
                codigoDepartamento = model.CodigoDepartamento,
                codigoProvincia = model.CodigoProvincia,
                codigoUbigeo = model.CodigoUbigeo,
                tipoDeporteId = model.TipoDeporteId
            });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var vmError = await ConstruirHomeVmAsync(model.Fecha, model.HoraInicio, model.HoraFin, model.CodigoDepartamento, model.CodigoProvincia, model.CodigoUbigeo, model.TipoDeporteId);
            return View("Index", vmError);
        }
    }

    [HttpGet]
    public async Task<IActionResult> UbigeoProvincias(string? codigoDepartamento)
    {
        var codigoDep = (codigoDepartamento ?? string.Empty).Trim();
        if (codigoDep.Length != 2) return Json(Array.Empty<object>());

        var data = await spService.UbigeoProvinciasListarAsync(codigoDep);
        return Json(data.Select(x => new { value = x.Value, text = x.Text }));
    }

    [HttpGet]
    public async Task<IActionResult> UbigeoDistritos(string? codigoProvincia)
    {
        var codigoProv = (codigoProvincia ?? string.Empty).Trim();
        if (codigoProv.Length != 4) return Json(Array.Empty<object>());

        var data = await spService.UbigeoDistritosListarAsync(codigoProv);
        return Json(data.Select(x => new { value = x.Value, text = x.Text }));
    }

    [HttpGet]
    public IActionResult ConsultarSolicitud()
    {
        return View(new SolicitudPublicaSeguimientoViewModel());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ConsultarSolicitud(SolicitudPublicaSeguimientoViewModel model)
    {
        if (!ModelState.IsValid) return View(model);

        var resultado = await spService.HomeConsultarSolicitudAsync(model.CodigoSolicitud.Trim(), model.Telefono.Trim());
        if (resultado is null)
        {
            model.Mensaje = "No se encontro una solicitud con ese codigo y telefono.";
            return View(model);
        }

        model.Resultado = resultado;
        return View(model);
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [HttpGet]
    public IActionResult SoftwareClubes()
    {
        var vm = new AltaClubSolicitudFormViewModel();
        AsignarCaptchaSoftwareClubes(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SoftwareClubes(AltaClubSolicitudFormViewModel model)
    {
        var captchaEsperado = HttpContext.Session.GetString(CaptchaSoftwareClubesSessionKey);
        if (string.IsNullOrWhiteSpace(captchaEsperado) ||
            !string.Equals(model.CaptchaCodigo?.Trim(), captchaEsperado, StringComparison.OrdinalIgnoreCase))
        {
            ModelState.AddModelError(nameof(model.CaptchaCodigo), "El codigo CAPTCHA no es valido.");
        }

        if (!ModelState.IsValid)
        {
            AsignarCaptchaSoftwareClubes(model);
            return View(model);
        }

        try
        {
            var correo = model.Correo.Trim();
            var existe = await userManager.FindByEmailAsync(correo);
            if (existe is not null)
            {
                ModelState.AddModelError(nameof(model.Correo), "Ya existe una cuenta con este correo.");
                AsignarCaptchaSoftwareClubes(model);
                return View(model);
            }

            var nuevoUsuario = new ApplicationUser
            {
                UserName = correo,
                Email = correo,
                Nombres = model.NombreContacto.Trim()
            };

            var resultadoCreacion = await userManager.CreateAsync(nuevoUsuario, model.Password);
            if (!resultadoCreacion.Succeeded)
            {
                foreach (var error in resultadoCreacion.Errors)
                {
                    ModelState.AddModelError(string.Empty, TraducirErrorIdentity(error.Code, error.Description));
                }

                AsignarCaptchaSoftwareClubes(model);
                return View(model);
            }

            try
            {
                var codigo = await spService.HomeRegistrarClubConPruebaAsync(model, nuevoUsuario.Id);
                await signInManager.SignInAsync(nuevoUsuario, isPersistent: false);

                var negocios = await spService.PanelListarNegociosUsuarioAsync(nuevoUsuario.Id);
                var negocioId = negocios.FirstOrDefault()?.NegocioId;

                TempData["MensajeSolicitudClub"] = $"Registro completado. Codigo: {codigo}. Tu prueba de 1 mes ya esta activa.";
                if (negocioId.HasValue)
                {
                    return RedirectToAction("Create", "Sedes", new { negocioId = negocioId.Value });
                }

                return RedirectToAction("Index", "Panel");
            }
            catch
            {
                await userManager.DeleteAsync(nuevoUsuario);
                throw;
            }
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            AsignarCaptchaSoftwareClubes(model);
            return View(model);
        }
    }

    private void AsignarCaptchaSoftwareClubes(AltaClubSolicitudFormViewModel model)
    {
        const string chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
        var captcha = new string(Enumerable.Range(0, 5)
            .Select(_ => chars[Random.Shared.Next(chars.Length)])
            .ToArray());

        HttpContext.Session.SetString(CaptchaSoftwareClubesSessionKey, captcha);
        model.CaptchaTexto = captcha;
        model.CaptchaCodigo = string.Empty;
    }

    private static string TraducirErrorIdentity(string code, string fallback)
    {
        return code switch
        {
            "PasswordRequiresNonAlphanumeric" => "La contraseña debe incluir al menos un símbolo (por ejemplo: !, @, #).",
            "PasswordRequiresLower" => "La contraseña debe incluir al menos una letra minúscula (a-z).",
            "PasswordRequiresUpper" => "La contraseña debe incluir al menos una letra mayúscula (A-Z).",
            "PasswordRequiresDigit" => "La contraseña debe incluir al menos un número (0-9).",
            "PasswordTooShort" => "La contraseña es muy corta. Usa al menos 6 caracteres.",
            "DuplicateEmail" => "Ya existe una cuenta registrada con este correo.",
            "DuplicateUserName" => "Ese correo/usuario ya está en uso.",
            "InvalidEmail" => "El correo ingresado no tiene un formato válido.",
            "InvalidUserName" => "El correo/usuario contiene caracteres no permitidos.",
            _ => fallback
        };
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    private async Task<HomeIndexViewModel> ConstruirHomeVmAsync(
        DateOnly? fecha,
        TimeOnly? horaInicio,
        TimeOnly? horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId)
    {
        var fechaConsulta = fecha ?? DateOnly.FromDateTime(DateTime.Today);
        var horaInicioConsulta = horaInicio ?? new TimeOnly(18, 0);
        var horaFinConsulta = horaFin ?? new TimeOnly(19, 0);
        var codigoDep = string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim();
        var codigoProv = string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim();
        var codigoDist = string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim();

        if (horaFinConsulta <= horaInicioConsulta)
            horaFinConsulta = horaInicioConsulta.AddHours(1);

        var sedes = await spService.HomeListarSedesAsync();
        var deportes = await spService.HomeListarTiposDeporteAsync();
        var espaciosDisponibles = await spService.HomeBuscarEspaciosDisponiblesAsync(fechaConsulta, horaInicioConsulta, horaFinConsulta, codigoDep, codigoProv, codigoDist, tipoDeporteId);
        var departamentos = await spService.UbigeoDepartamentosListarAsync();
        var provincias = !string.IsNullOrWhiteSpace(codigoDep) && codigoDep.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(codigoDep)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
        var distritos = !string.IsNullOrWhiteSpace(codigoProv) && codigoProv.Length == 4
            ? await spService.UbigeoDistritosListarAsync(codigoProv)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();

        return new HomeIndexViewModel
        {
            Fecha = fechaConsulta,
            HoraInicio = horaInicioConsulta,
            HoraFin = horaFinConsulta,
            CodigoDepartamento = codigoDep,
            CodigoProvincia = codigoProv,
            CodigoUbigeo = codigoDist,
            TipoDeporteId = tipoDeporteId,
            DepartamentosUbigeo = departamentos,
            ProvinciasUbigeo = provincias,
            DistritosUbigeo = distritos,
            Sedes = sedes,
            TiposDeporte = deportes,
            Disponibles = espaciosDisponibles
        };
    }
}
