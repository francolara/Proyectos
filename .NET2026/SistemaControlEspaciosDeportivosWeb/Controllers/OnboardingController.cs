using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class OnboardingController(ISportCenterStoredProcedureService spService) : Controller
{
    public async Task<IActionResult> Index(int? negocioId, byte? paso = null)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId);
        if (!resolvedNegocioId.HasValue)
            return Forbid();

        ViewData["AdminShell"] = true;
        ViewData["AdminNegocioId"] = resolvedNegocioId.Value;

        var checklist = await spService.OnboardingChecklistValidarAsync(resolvedNegocioId.Value);
        var pasoPendiente = ResolverPasoPendiente(checklist);
        var config = await spService.ConfiguracionClubObtenerAsync(resolvedNegocioId.Value);
        var pasoActual = checklist.ChecklistCompleto
            ? (byte)5
            : (paso.HasValue && paso.Value >= 1 && paso.Value <= 5 ? paso.Value : pasoPendiente);
        var vm = await ConstruirDashboardAsync(resolvedNegocioId.Value, config?.NombreComercial ?? string.Empty, config?.LogoUrl, checklist, pasoActual);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> IrPaso(int negocioId, byte paso)
    {
        var checklist = await spService.OnboardingChecklistValidarAsync(negocioId);
        var pasoPendiente = ResolverPasoPendiente(checklist);
        var pasoDestino = paso < pasoPendiente ? pasoPendiente : paso;
        if (pasoDestino < 1) pasoDestino = 1;
        if (pasoDestino > 5) pasoDestino = 5;
        return RedirectToAction(nameof(Index), new { negocioId, paso = pasoDestino });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarConfiguracion(OnboardingConfiguracionFormViewModel form)
    {
        var model = await spService.ConfiguracionClubObtenerAsync(form.NegocioId) ?? new ConfiguracionClubViewModel { NegocioId = form.NegocioId, Id = form.NegocioId };
        model.NombreComercial = form.NombreComercial?.Trim() ?? string.Empty;
        model.TipoDocumento = form.TipoDocumento?.Trim() ?? "1";
        model.NumeroDocumento = string.IsNullOrWhiteSpace(form.NumeroDocumento) ? null : form.NumeroDocumento.Trim();
        model.MonedaId = form.MonedaId;

        var ok = await spService.ConfiguracionClubActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (ok)
        {
            TempData["OnboardingOk"] = "Configuracion guardada. ContinÃºa con Maestros.";
            return RedirectToAction(nameof(Index), new { negocioId = form.NegocioId, paso = 2 });
        }

        TempData["OnboardingInfo"] = "No se pudo guardar la configuracion.";
        return RedirectToAction(nameof(Index), new { negocioId = form.NegocioId, paso = 1 });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarMaestros(OnboardingMaestrosFormViewModel form)
    {
        var usuario = User.Identity?.Name ?? "sistema";
        if (form.TipoDeporteSuperId.HasValue && form.TipoDeporteSuperId.Value > 0)
            await spService.MaestrosTiposDeporteCrearAsync(form.NegocioId, form.TipoDeporteSuperId.Value, true, usuario);
        if (form.TipoSueloSuperId.HasValue && form.TipoSueloSuperId.Value > 0)
            await spService.MaestrosTiposSueloCrearAsync(form.NegocioId, form.TipoSueloSuperId.Value, true, usuario);
        if (form.MonedaSuperId.HasValue && form.MonedaSuperId.Value > 0)
            await spService.MaestrosMonedasCrearAsync(form.NegocioId, form.MonedaSuperId.Value, true, usuario);
        if (!string.IsNullOrWhiteSpace(form.FormaPagoNombre))
            await spService.MaestrosFormasPagoCrearAsync(form.NegocioId, form.FormaPagoNombre.Trim(), true, usuario);
        if (!string.IsNullOrWhiteSpace(form.CodigoSunatDocumento))
            await spService.MaestrosTiposDocumentoComprobanteCrearAsync(form.NegocioId, form.CodigoSunatDocumento.Trim(), true, usuario);
        if (!string.IsNullOrWhiteSpace(form.CodigoSunatDocumento) && !string.IsNullOrWhiteSpace(form.SerieDocumento))
            await spService.ConfiguracionSeriesDocumentoGuardarAsync(form.NegocioId, form.CodigoSunatDocumento.Trim(), form.SerieDocumento.Trim().ToUpperInvariant(), true, usuario);

        var checklist = await spService.OnboardingChecklistValidarAsync(form.NegocioId);
        var paso = ResolverPasoPendiente(checklist);
        TempData["OnboardingOk"] = "Datos de Maestros guardados.";
        return RedirectToAction(nameof(Index), new { negocioId = form.NegocioId, paso });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarSede(OnboardingSedeFormViewModel form)
    {
        var sede = new SedeFormViewModel
        {
            NegocioId = form.NegocioId,
            Nombre = form.Nombre.Trim(),
            Direccion = form.Direccion.Trim(),
            CodigoUbigeo = form.CodigoUbigeo.Trim(),
            HoraApertura = form.HoraApertura,
            HoraCierre = form.HoraCierre,
            Activo = true,
            ServiciosSeleccionados = form.ServiciosSeleccionados,
            NotificacionesActivas = true
        };
        await spService.SedesCrearAsync(sede, User.Identity?.Name ?? "sistema");
        var checklist = await spService.OnboardingChecklistValidarAsync(form.NegocioId);
        var paso = ResolverPasoPendiente(checklist);
        TempData["OnboardingOk"] = "Sede registrada.";
        return RedirectToAction(nameof(Index), new { negocioId = form.NegocioId, paso });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarEspacio(OnboardingEspacioFormViewModel form)
    {
        var espacio = new EspacioFormViewModel
        {
            NegocioId = form.NegocioId,
            SedeId = form.SedeId,
            TipoDeporteId = form.TipoDeporteId,
            TipoSueloId = form.TipoSueloId,
            Codigo = form.Codigo.Trim(),
            Nombre = form.Nombre.Trim(),
            Capacidad = form.Capacidad,
            Estado = EstadoEspacioDeportivo.Activo,
            Tarifas =
            [
                new EspacioTarifaRangoViewModel
                {
                    DiaSemana = 1,
                    HoraInicio = new TimeOnly(8,0),
                    HoraFin = new TimeOnly(9,0),
                    Precio = form.PrecioBase
                }
            ]
        };
        await spService.EspaciosCrearAsync(espacio, User.Identity?.Name ?? "sistema");
        var checklist = await spService.OnboardingChecklistValidarAsync(form.NegocioId);
        var paso = ResolverPasoPendiente(checklist);
        TempData["OnboardingOk"] = "Espacio registrado.";
        return RedirectToAction(nameof(Index), new { negocioId = form.NegocioId, paso });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Finalizar(int negocioId)
    {
        var checklist = await spService.OnboardingChecklistValidarAsync(negocioId);
        if (!checklist.ChecklistCompleto)
        {
            var pasoPendiente = ResolverPasoPendiente(checklist);
            TempData["OnboardingInfo"] = "Aun hay requisitos pendientes de onboarding.";
            return await RedirigirPasoAsync(negocioId, pasoPendiente);
        }
        TempData["OnboardingOk"] = "Onboarding completado correctamente.";
        return RedirectToAction("Index", "Panel", new { negocioId });
    }

    private async Task<IActionResult> RedirigirPasoAsync(int negocioId, byte pasoPendiente)
    {
        var primeraSede = (await spService.SedesListarAsync(negocioId)).FirstOrDefault();
        var primerEspacio = (await spService.EspaciosListarAsync(negocioId)).FirstOrDefault();

        return pasoPendiente switch
        {
            1 => RedirectToAction("Index", "Configuracion", new { negocioId }),
            2 => RedirectToAction("Index", "Maestros", new { negocioId }),
            3 => primeraSede is null
                ? RedirectToAction("Create", "Sedes", new { negocioId })
                : RedirectToAction("Edit", "Sedes", new { id = primeraSede.Id, negocioId }),
            4 => primerEspacio is null
                ? RedirectToAction("Create", "Espacios", new { negocioId })
                : RedirectToAction("Edit", "Espacios", new { id = primerEspacio.Id, negocioId }),
            _ => RedirectToAction("Index", "Panel", new { negocioId })
        };
    }

    private static byte ResolverPasoPendiente(OnboardingChecklistViewModel checklist)
    {
        var pasoConfiguracionOk = checklist.ConfigNombreComercialOk
                                  && checklist.ConfigTipoDocumentoOk
                                  && checklist.ConfigMonedaOk
                                  && checklist.ConfigCpeCondicionesOk;
        if (!pasoConfiguracionOk) return 1;

        var pasoMaestrosOk = checklist.MaestroTipoDeporteOk
                              && checklist.MaestroTipoSueloOk
                              && checklist.MaestroFormaPagoOk
                              && checklist.MaestroMonedaOk;
        if (!pasoMaestrosOk) return 2;

        if (!checklist.SedeMinimaOk) return 3;
        if (!checklist.EspacioMinimoOk) return 4;
        return 5;
    }

    private async Task<OnboardingDashboardViewModel> ConstruirDashboardAsync(
        int negocioId,
        string negocioNombre,
        string? logoUrl,
        OnboardingChecklistViewModel checklist,
        byte pasoActual)
    {
        var sedesExistentes = await spService.SedesListarAsync(negocioId);
        var primeraSede = sedesExistentes.FirstOrDefault();
        var espaciosExistentes = await spService.EspaciosListarAsync(negocioId);
        var primerEspacio = espaciosExistentes.FirstOrDefault();

        var pasos = new List<OnboardingPasoItemViewModel>
        {
            new()
            {
                Paso = 1,
                Titulo = "Configuracion",
                Descripcion = "Completa razon social, documento, direccion, ubigeo, IGV y reglas de reserva.",
                Completado = checklist.ConfigNombreComercialOk
                             && checklist.ConfigTipoDocumentoOk
                             && checklist.ConfigMonedaOk
                             && checklist.ConfigCpeCondicionesOk,
                UrlAccion = "/Configuracion/Index?negocioId=" + negocioId
            },
            new()
            {
                Paso = 2,
                Titulo = "Maestros",
                Descripcion = "Activa deporte, suelo, moneda y forma de pago.",
                Completado = checklist.MaestroTipoDeporteOk
                             && checklist.MaestroTipoSueloOk
                             && checklist.MaestroFormaPagoOk
                             && checklist.MaestroMonedaOk,
                UrlAccion = "/Maestros/Index?negocioId=" + negocioId
            },
            new()
            {
                Paso = 3,
                Titulo = "Sedes",
                Descripcion = "Registra sede valida con horario, servicios, notificaciones, correo y WhatsApp.",
                Completado = checklist.SedeMinimaOk,
                UrlAccion = primeraSede is null
                    ? "/Sedes/Create?negocioId=" + negocioId
                    : $"/Sedes/Edit/{primeraSede.Id}?negocioId={negocioId}"
            },
            new()
            {
                Paso = 4,
                Titulo = "Espacios",
                Descripcion = "Registra al menos un espacio activo con tarifa valida.",
                Completado = checklist.EspacioMinimoOk,
                UrlAccion = primerEspacio is null
                    ? "/Espacios/Create?negocioId=" + negocioId
                    : $"/Espacios/Edit/{primerEspacio.Id}?negocioId={negocioId}"
            },
            new()
            {
                Paso = 5,
                Titulo = "Resumen",
                Descripcion = "Revisa checklist y finaliza onboarding.",
                Completado = checklist.ChecklistCompleto,
                UrlAccion = "/Onboarding/Index?negocioId=" + negocioId
            }
        };

        foreach (var paso in pasos)
            paso.EsActual = paso.Paso == pasoActual;

        var config = await spService.ConfiguracionClubObtenerAsync(negocioId) ?? new ConfiguracionClubViewModel { NegocioId = negocioId };
        var servicios = await spService.SedesComboServiciosAsync();
        var sedes = await spService.EspaciosComboSedesAsync(negocioId);
        var deportes = await spService.EspaciosComboTiposDeporteAsync(negocioId);
        var suelos = await spService.EspaciosComboTiposSueloAsync(negocioId);

        return new OnboardingDashboardViewModel
        {
            NegocioId = negocioId,
            NegocioNombre = negocioNombre,
            LogoUrl = logoUrl,
            PasoActual = pasoActual,
            PasosCompletados = pasos.Count(x => x.Completado),
            TotalPasos = 5,
            ChecklistCompleto = checklist.ChecklistCompleto,
            Checklist = checklist,
            Pasos = pasos,
            ConfiguracionForm = new OnboardingConfiguracionFormViewModel
            {
                NegocioId = negocioId,
                NombreComercial = config.NombreComercial,
                TipoDocumento = config.TipoDocumento,
                NumeroDocumento = config.NumeroDocumento,
                MonedaId = config.MonedaId,
                TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync(),
                Monedas = await spService.ConfiguracionClubComboMonedasAsync(negocioId)
            },
            MaestrosForm = new OnboardingMaestrosFormViewModel
            {
                NegocioId = negocioId,
                TiposDeporteSuper = await spService.MaestrosTiposDeporteSuperListarAsync(),
                TiposSueloSuper = await spService.MaestrosTiposSueloSuperListarAsync(),
                MonedasSuper = await spService.MaestrosMonedasSuperListarAsync(),
                TiposDocumentoComprobanteSuper = await spService.MaestrosTiposDocumentoComprobanteSuperListarAsync()
            },
            SedeForm = new OnboardingSedeFormViewModel
            {
                NegocioId = negocioId,
                ServiciosDisponibles = servicios
            },
            EspacioForm = new OnboardingEspacioFormViewModel
            {
                NegocioId = negocioId,
                Sedes = sedes,
                TiposDeporte = deportes,
                TiposSuelo = suelos
            }
        };
    }

    private async Task<int?> ResolverNegocioIdAsync(int? negocioId)
    {
        if (negocioId.HasValue && negocioId.Value > 0)
            return negocioId.Value;

        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId))
            return null;

        var membresias = await spService.PanelListarNegociosUsuarioAsync(usuarioId);
        return membresias.FirstOrDefault()?.NegocioId;
    }
}

