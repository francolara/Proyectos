using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Diagnostics;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class HomeController(ISportCenterStoredProcedureService spService, INotificacionEmailService notificacionEmailService) : Controller
{
    public async Task<IActionResult> Index(
        DateOnly? fecha,
        TimeOnly? horaInicio,
        TimeOnly? horaFin,
        int? sedeId,
        int? tipoDeporteId)
    {
        var vm = await ConstruirHomeVmAsync(fecha, horaInicio, horaFin, sedeId, tipoDeporteId);
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
            var vmError = await ConstruirHomeVmAsync(model.Fecha, model.HoraInicio, model.HoraFin, model.SedeId, model.TipoDeporteId);
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
                sedeId = model.SedeId,
                tipoDeporteId = model.TipoDeporteId
            });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var vmError = await ConstruirHomeVmAsync(model.Fecha, model.HoraInicio, model.HoraFin, model.SedeId, model.TipoDeporteId);
            return View("Index", vmError);
        }
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

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    private async Task<HomeIndexViewModel> ConstruirHomeVmAsync(
        DateOnly? fecha,
        TimeOnly? horaInicio,
        TimeOnly? horaFin,
        int? sedeId,
        int? tipoDeporteId)
    {
        var fechaConsulta = fecha ?? DateOnly.FromDateTime(DateTime.Today);
        var horaInicioConsulta = horaInicio ?? new TimeOnly(18, 0);
        var horaFinConsulta = horaFin ?? new TimeOnly(19, 0);

        if (horaFinConsulta <= horaInicioConsulta)
            horaFinConsulta = horaInicioConsulta.AddHours(1);

        var sedes = await spService.HomeListarSedesAsync();
        var deportes = await spService.HomeListarTiposDeporteAsync();
        var espaciosDisponibles = await spService.HomeBuscarEspaciosDisponiblesAsync(fechaConsulta, horaInicioConsulta, horaFinConsulta, sedeId, tipoDeporteId);

        return new HomeIndexViewModel
        {
            Fecha = fechaConsulta,
            HoraInicio = horaInicioConsulta,
            HoraFin = horaFinConsulta,
            SedeId = sedeId,
            TipoDeporteId = tipoDeporteId,
            Sedes = sedes,
            TiposDeporte = deportes,
            Disponibles = espaciosDisponibles
        };
    }
}
