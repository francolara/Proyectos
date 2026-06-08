using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Diagnostics;
using System.Globalization;
using System.Security.Claims;
using Microsoft.AspNetCore.Mvc.ModelBinding;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class HomeController(
    ISportCenterStoredProcedureService spService,
    IReservationEmailNotificationService reservationEmailNotificationService,
    UserManager<ApplicationUser> userManager,
    ILogger<HomeController> logger) : Controller
{
    private const int DuracionReservaPublicaMinutos = 60;
    private static readonly TimeOnly HoraInicioCierreMedianoche = new(23, 0);
    private static readonly TimeOnly HoraFinCierreMedianoche = new(23, 59);
    
    private static (DateOnly Fecha, TimeOnly HoraInicio, TimeOnly HoraFin) ObtenerRangoSugeridoBusqueda(DateTime ahoraLocal)
    {
        var ahoraSinSegundos = new DateTime(
            ahoraLocal.Year,
            ahoraLocal.Month,
            ahoraLocal.Day,
            ahoraLocal.Hour,
            ahoraLocal.Minute,
            0,
            ahoraLocal.Kind);

        var minutosRestantes = 30 - (ahoraSinSegundos.Minute % 30);
        if (minutosRestantes == 30)
            minutosRestantes = 0;

        var baseHora = ahoraSinSegundos.AddMinutes(minutosRestantes);
        if (baseHora <= ahoraSinSegundos)
            baseHora = baseHora.AddMinutes(30);

        var horaInicio = new TimeOnly(baseHora.Hour, baseHora.Minute);
        var fecha = DateOnly.FromDateTime(baseHora.Date);

        // Evita sugerir horarios de madrugada como predeterminados en el portal.
        if (horaInicio < new TimeOnly(6, 0))
        {
            horaInicio = new TimeOnly(18, 0);
            fecha = DateOnly.FromDateTime(ahoraLocal.Date);
        }

        // Si es muy tarde, propone el siguiente dia en franja comercial.
        if (horaInicio >= new TimeOnly(23, 0))
        {
            fecha = fecha.AddDays(1);
            horaInicio = new TimeOnly(18, 0);
        }

        var horaFin = horaInicio.AddHours(1);
        if (horaFin <= horaInicio)
            horaFin = new TimeOnly(23, 59);

        return (fecha, horaInicio, horaFin);
    }
    public async Task<IActionResult> Index(
        DateOnly? fecha,
        TimeOnly? horaInicio,
        TimeOnly? horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId,
        int? negocioId,
        bool buscarCercaDeMi = false,
        decimal? latitudUsuario = null,
        decimal? longitudUsuario = null,
        decimal? radioKm = null,
        bool omitirFechaHorario = true,
        int pagina = 1)
    {
        ViewData["PublicFullWidth"] = true;
        ViewData["HideDefaultFooter"] = true;
        var vm = await ConstruirHomeVmAsync(
            fecha,
            horaInicio,
            horaFin,
            codigoDepartamento,
            codigoProvincia,
            codigoUbigeo,
            tipoDeporteId,
            negocioId,
            buscarCercaDeMi,
            latitudUsuario,
            longitudUsuario,
            radioKm,
            omitirFechaHorario,
            pagina);
        ViewData["MostrarPromosNav"] = vm.PopupPromociones.Count > 0;
        vm.MensajeSolicitud = TempData["MensajeSolicitud"]?.ToString();
        return View(vm);
    }

    [HttpGet]
    public IActionResult Faq()
    {
        ViewData["PublicFullWidth"] = true;
        return View();
    }

    [HttpGet]
    public async Task<IActionResult> Reservar(
        int espacioDeportivoId,
        DateOnly? fecha,
        TimeOnly? horaInicio,
        TimeOnly? horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId,
        int? negocioId,
        bool omitirFechaHorario = false)
    {
        ViewData["PublicFullWidth"] = true;

        var fechaConsulta = fecha ?? DateOnly.FromDateTime(DateTime.Today);
        var horaInicioConsulta = horaInicio ?? new TimeOnly(18, 0);
        var horaFinConsulta = horaFin ?? horaInicioConsulta.AddHours(1);
        if (horaFinConsulta <= horaInicioConsulta)
            horaFinConsulta = new TimeOnly(23, 59);

        var codigoDep = string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim();
        var codigoProv = string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim();
        var codigoDist = string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim();
        var vm = await ConstruirReservaVmAsync(
            espacioDeportivoId,
            fechaConsulta,
            horaInicioConsulta,
            horaFinConsulta,
            codigoDep,
            codigoProv,
            codigoDist,
            tipoDeporteId,
            negocioId,
            omitirFechaHorario: omitirFechaHorario);

        if (vm is null)
        {
            TempData["MensajeSolicitud"] = "El espacio ya no esta disponible para el horario seleccionado.";
            return RedirectToAction(nameof(Index), new
            {
                fecha = fechaConsulta,
                horaInicio = horaInicioConsulta,
                horaFin = horaFinConsulta,
                codigoDepartamento = codigoDep,
                codigoProvincia = codigoProv,
                codigoUbigeo = codigoDist,
                tipoDeporteId,
                negocioId,
                omitirFechaHorario,
                pagina = 1
            });
        }

        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> ReservarCalendarioEventos(
        int negocioId,
        int espacioDeportivoId,
        DateTime? start,
        DateTime? end)
    {
        var desde = DateOnly.FromDateTime((start ?? DateTime.Today).Date);
        var hasta = DateOnly.FromDateTime((end ?? DateTime.Today).Date);
        if (hasta < desde)
            hasta = desde;

        var eventos = await spService.ReservasCalendarioEventosAsync(
            negocioId,
            desde,
            hasta,
            sedeId: null,
            espacioDeportivoId: espacioDeportivoId,
            estado: null);

        var payload = new List<object>();
        foreach (var e in eventos.Where(e =>
                     string.Equals(e.TipoEvento, "RESERVA", StringComparison.OrdinalIgnoreCase)
                     || string.Equals(e.TipoEvento, "RESERVA_COMPARTIDA", StringComparison.OrdinalIgnoreCase)
                     || string.Equals(e.TipoEvento, "BLOQUEO", StringComparison.OrdinalIgnoreCase)
                     || string.Equals(e.TipoEvento, "NO_ATENCION", StringComparison.OrdinalIgnoreCase)))
        {
            var inicio = e.Fecha.ToDateTime(e.HoraInicio);
            var fin = e.Fecha.ToDateTime(e.HoraFin);
            if (fin <= inicio) fin = inicio.AddMinutes(30);

            if (string.Equals(e.TipoEvento, "RESERVA", StringComparison.OrdinalIgnoreCase))
            {
                var estadoUi = e.Estado switch
                {
                    1 => "reservada",
                    2 or 3 or 4 => "confirmada",
                    _ => string.Empty
                };

                if (string.IsNullOrWhiteSpace(estadoUi))
                    continue;

                payload.Add(new
                {
                    id = $"r-{e.Id}",
                    title = ExtraerTituloPublico(e.Titulo),
                    start = inicio.ToString("yyyy-MM-ddTHH:mm:ss"),
                    end = fin.ToString("yyyy-MM-ddTHH:mm:ss"),
                    color = estadoUi == "reservada" ? "#f59f00" : "#2563eb",
                    classNames = new[] { "sc-public-event", $"is-{estadoUi}" },
                    extendedProps = new
                    {
                        tipo = "RESERVA",
                        estado = estadoUi
                    }
                });
                continue;
            }

            payload.Add(new
            {
                id = $"b-{e.Id}",
                title = ExtraerTituloPublico(e.Titulo),
                start = inicio.ToString("yyyy-MM-ddTHH:mm:ss"),
                end = fin.ToString("yyyy-MM-ddTHH:mm:ss"),
                color = "#334155",
                display = "background",
                classNames = new[]
                {
                    string.Equals(e.TipoEvento, "NO_ATENCION", StringComparison.OrdinalIgnoreCase) ? "sc-no-atencion-bg" : "sc-bloqueo-bg",
                    "is-bloqueado"
                },
                extendedProps = new
                {
                    tipo = e.TipoEvento,
                    estado = "bloqueado",
                    motivo = ExtraerTituloPublico(e.Titulo)
                }
            });
        }

        return Json(payload);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CrearReservaPublica(SolicitudReservaPublicaFormViewModel model)
    {
        return await ProcesarCreacionReservaPublicaAsync(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ActionName("SolicitarReservaPublica")]
    public async Task<IActionResult> SolicitarReservaPublicaLegacy(SolicitudReservaPublicaFormViewModel model)
    {
        return await ProcesarCreacionReservaPublicaAsync(model);
    }

    private async Task<IActionResult> ProcesarCreacionReservaPublicaAsync(SolicitudReservaPublicaFormViewModel model)
    {
        logger.LogInformation(
            "Inicio crear reserva publica. EspacioDeportivoId={EspacioDeportivoId}, NegocioId={NegocioId}, Fecha={Fecha}, HoraInicio={HoraInicio}, HoraFin={HoraFin}.",
            model.EspacioDeportivoId,
            model.NegocioId,
            model.Fecha,
            model.HoraInicio,
            model.HoraFin);

        ViewData["PublicFullWidth"] = true;
        RemoverModelStatePorPrefijo(ModelState, "TiposDocumentoIdentidad");
        ModelState.Remove(nameof(model.OmitirFechaHorario));
        ModelState.Remove("OmitirFechaHorario");
        var omitirFechaHorario = model.OmitirFechaHorario ?? false;
        if (User.Identity?.IsAuthenticated == true)
            model.UsuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);

        model.TipoDocumento = string.IsNullOrWhiteSpace(model.TipoDocumento) ? "0" : model.TipoDocumento.Trim();
        model.NumeroDocumento = string.IsNullOrWhiteSpace(model.NumeroDocumento) ? null : model.NumeroDocumento.Trim();
        model.Nombres = (model.Nombres ?? string.Empty).Trim();
        model.Apellidos = (model.Apellidos ?? string.Empty).Trim();
        model.NombreEquipo = string.IsNullOrWhiteSpace(model.NombreEquipo) ? null : model.NombreEquipo.Trim();
        model.Telefono = string.IsNullOrWhiteSpace(model.Telefono) ? null : model.Telefono.Trim();
        model.Correo = string.IsNullOrWhiteSpace(model.Correo) ? null : model.Correo.Trim();
        model.Comentario = string.IsNullOrWhiteSpace(model.Comentario) ? null : model.Comentario.Trim();
        model.CodigoCupon = string.IsNullOrWhiteSpace(model.CodigoCupon) ? null : model.CodigoCupon.Trim().ToUpperInvariant();
        model.HoraFin = NormalizarHoraFinReservaPublica(model.HoraInicio, model.HoraFin);

        if (model.HoraFin <= model.HoraInicio)
            ModelState.AddModelError(string.Empty, "La hora fin debe ser mayor que la hora inicio.");
        else if (!EsDuracionReservaPublicaValida(model.HoraInicio, model.HoraFin))
            ModelState.AddModelError(string.Empty, "La reserva publica solo permite bloques de 1 hora (o de 23:00 a 23:59).");

        if (!ModelState.IsValid)
        {
            logger.LogWarning("Reserva publica invalida. Detalle: {Detalle}",
                string.Join(" | ", ModelState
                    .Where(x => x.Value?.Errors?.Count > 0)
                    .SelectMany(x => x.Value!.Errors.Select(e => $"{(string.IsNullOrWhiteSpace(x.Key) ? "<sin-campo>" : x.Key)}: {e.ErrorMessage}"))));
            var detalleErrores = ObtenerDetalleErroresValidacion(ModelState);
            if (!string.IsNullOrWhiteSpace(detalleErrores))
                ModelState.AddModelError(string.Empty, "Campos pendientes: " + detalleErrores);

            var vmError = await ConstruirReservaVmAsync(
                model.EspacioDeportivoId,
                model.Fecha,
                model.HoraInicio,
                model.HoraFin,
                model.CodigoDepartamento,
                model.CodigoProvincia,
                model.CodigoUbigeo,
                model.TipoDeporteId,
                model.NegocioId,
                omitirFechaHorario: omitirFechaHorario,
                formBase: model);

            if (vmError is null)
            {
                TempData["MensajeSolicitud"] = "El espacio ya no esta disponible para el horario seleccionado.";
                return RedirectToAction(nameof(Index), new
                {
                    fecha = model.Fecha,
                    horaInicio = model.HoraInicio,
                    horaFin = model.HoraFin,
                    codigoDepartamento = model.CodigoDepartamento,
                    codigoProvincia = model.CodigoProvincia,
                    codigoUbigeo = model.CodigoUbigeo,
                    tipoDeporteId = model.TipoDeporteId,
                    negocioId = model.NegocioId,
                    omitirFechaHorario = omitirFechaHorario,
                    pagina = 1
                });
            }

            return View("Reservar", vmError);
        }

        try
        {
            if (!string.IsNullOrWhiteSpace(model.UsuarioId))
            {
                try
                {
                    await spService.UsuariosPublicosGuardarPerfilAsync(new UsuarioPublicoPerfilViewModel
                    {
                        UsuarioId = model.UsuarioId,
                        TipoDocumento = model.TipoDocumento,
                        NumeroDocumento = model.NumeroDocumento,
                        Nombres = model.Nombres,
                        Apellidos = model.Apellidos,
                        NombreEquipo = model.NombreEquipo,
                        Telefono = model.Telefono,
                        Correo = model.Correo
                    }, User.Identity?.Name ?? "portal-web");
                }
                catch
                {
                    // El perfil publico no debe bloquear la reserva.
                }
            }

            var reservaId = await spService.HomeSolicitarReservaPublicaAsync(model);
            await reservationEmailNotificationService.NotifyPublicReservationCreatedAsync(null, reservaId);
            logger.LogInformation(
                "Reserva publica creada con exito. ReservaId={ReservaId}.",
                reservaId);
            TempData["MensajeSolicitud"] = $"Reserva registrada correctamente. Codigo: R-{reservaId:D6}.";
            return RedirectToAction(nameof(Index), new
            {
                fecha = model.Fecha,
                horaInicio = model.HoraInicio,
                horaFin = model.HoraFin,
                codigoDepartamento = model.CodigoDepartamento,
                codigoProvincia = model.CodigoProvincia,
                codigoUbigeo = model.CodigoUbigeo,
                tipoDeporteId = model.TipoDeporteId,
                negocioId = model.NegocioId,
                omitirFechaHorario = omitirFechaHorario,
                pagina = 1
            });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var vmError = await ConstruirReservaVmAsync(
                model.EspacioDeportivoId,
                model.Fecha,
                model.HoraInicio,
                model.HoraFin,
                model.CodigoDepartamento,
                model.CodigoProvincia,
                model.CodigoUbigeo,
                model.TipoDeporteId,
                model.NegocioId,
                omitirFechaHorario: omitirFechaHorario,
                formBase: model);

            if (vmError is null)
            {
                TempData["MensajeSolicitud"] = "El espacio ya no esta disponible para el horario seleccionado.";
                return RedirectToAction(nameof(Index), new
                {
                    fecha = model.Fecha,
                    horaInicio = model.HoraInicio,
                    horaFin = model.HoraFin,
                    codigoDepartamento = model.CodigoDepartamento,
                    codigoProvincia = model.CodigoProvincia,
                    codigoUbigeo = model.CodigoUbigeo,
                    tipoDeporteId = model.TipoDeporteId,
                    negocioId = model.NegocioId,
                    omitirFechaHorario = omitirFechaHorario,
                    pagina = 1
                });
            }

            return View("Reservar", vmError);
        }
    }

    [HttpGet]
    public async Task<IActionResult> CotizarReservaPublica(
        int negocioId,
        int espacioDeportivoId,
        string fecha,
        string horaInicio,
        string horaFin,
        string? codigoCupon = null)
    {
        if (!DateOnly.TryParseExact(fecha, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var fechaParsed) ||
            !TimeOnly.TryParseExact(horaInicio, "HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var horaInicioParsed) ||
            !TimeOnly.TryParseExact(horaFin, "HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var horaFinParsed))
        {
            return Json(new { ok = false, mensaje = "Fecha u hora invalidas." });
        }

        horaFinParsed = NormalizarHoraFinReservaPublica(horaInicioParsed, horaFinParsed);

        if (horaFinParsed <= horaInicioParsed)
            return Json(new { ok = false, mensaje = "La hora fin debe ser mayor que la hora inicio." });
        if (!EsDuracionReservaPublicaValida(horaInicioParsed, horaFinParsed))
            return Json(new { ok = false, mensaje = "La reserva publica solo permite bloques de 1 hora (o de 23:00 a 23:59)." });

        try
        {
            var cotizacion = await spService.ReservasCotizarAsync(negocioId, espacioDeportivoId, fechaParsed, horaInicioParsed, horaFinParsed);
            var montoDescuentoCupon = 0m;
            var montoFinalConCupon = cotizacion.PrecioFinal;
            var cuponAplicado = string.Empty;
            var mensajeCupon = string.Empty;
            if (!string.IsNullOrWhiteSpace(codigoCupon))
            {
                var validacion = await spService.CuponesValidarAsync(negocioId, null, espacioDeportivoId, codigoCupon, cotizacion.PrecioFinal);
                if (validacion.EsValido)
                {
                    montoDescuentoCupon = validacion.MontoDescuento;
                    montoFinalConCupon = validacion.MontoFinal;
                    cuponAplicado = validacion.CodigoCupon;
                }

                mensajeCupon = validacion.Mensaje;
            }
            return Json(new
            {
                ok = true,
                cotizacion = new
                {
                    mensaje = cotizacion.Mensaje,
                    precioBase = cotizacion.PrecioBase,
                    descuentoPct = cotizacion.DescuentoPct,
                    precioFinal = cotizacion.PrecioFinal,
                    montoDescuentoCupon,
                    montoFinalConCupon,
                    cuponAplicado,
                    mensajeCupon,
                    monedaSimbolo = cotizacion.MonedaSimbolo,
                    politicaConfirmacionPago = cotizacion.PoliticaConfirmacionPago,
                    porcentajeAdelantoMinimo = cotizacion.PorcentajeAdelantoMinimo
                }
            });
        }
        catch (Exception ex)
        {
            return Json(new { ok = false, mensaje = ex.Message });
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
    public async Task<IActionResult> NegociosPorUbigeo(string? codigoDepartamento, string? codigoProvincia, string? codigoUbigeo)
    {
        var codigoDep = string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim();
        var codigoProv = string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim();
        var codigoDist = string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim();

        var sedes = await spService.HomeListarSedesAsync();
        var negocios = ConstruirNegociosFiltradosPorUbigeo(sedes, codigoDep, codigoProv, codigoDist);
        return Json(negocios.Select(x => new { value = x.Value, text = x.Text }));
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
        return RedirectToPage("/Account/Register", new { area = "Identity", TipoRegistro = "club" });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public IActionResult SoftwareClubes(AltaClubSolicitudFormViewModel model)
    {
        return RedirectToPage("/Account/Register", new { area = "Identity", TipoRegistro = "club" });
    }

    private static void RemoverModelStatePorPrefijo(ModelStateDictionary modelState, string prefijo)
    {
        var keys = modelState.Keys
            .Where(k => string.Equals(k, prefijo, StringComparison.OrdinalIgnoreCase)
                     || k.StartsWith(prefijo + ".", StringComparison.OrdinalIgnoreCase)
                     || k.StartsWith(prefijo + "[", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var key in keys)
            modelState.Remove(key);
    }

    private static string ObtenerDetalleErroresValidacion(ModelStateDictionary modelState)
    {
        var campos = new List<string>();
        foreach (var entry in modelState)
        {
            if (entry.Value.Errors.Count == 0) continue;

            var key = entry.Key;
            if (string.IsNullOrWhiteSpace(key))
            {
                var texto = entry.Value.Errors
                    .Select(e => (e.ErrorMessage ?? string.Empty).Trim())
                    .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
                if (!string.IsNullOrWhiteSpace(texto))
                    campos.Add("<sin-campo>: " + texto);
                continue;
            }

            var campo = key.Split('.').Last();
            if (string.Equals(campo, "OmitirFechaHorario", StringComparison.OrdinalIgnoreCase))
                continue;
            campo = campo.Replace("HoraInicio", "Hora inicio")
                         .Replace("HoraFin", "Hora fin")
                         .Replace("NumeroDocumento", "Numero de documento")
                         .Replace("TipoDocumento", "Tipo de documento")
                         .Replace("NombreEquipo", "Nombre de equipo")
                         .Replace("EspacioDeportivoId", "Espacio deportivo")
                         .Replace("CodigoUbigeo", "Distrito");
            campos.Add(campo);
        }

        return string.Join(", ", campos.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    private static string ExtraerTituloPublico(string? tituloOriginal)
    {
        var raw = (tituloOriginal ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(raw))
            return "Reservada";

        var idxGuion = raw.IndexOf(" - ", StringComparison.Ordinal);
        var baseTitulo = idxGuion >= 0 ? raw[(idxGuion + 3)..].Trim() : raw;
        var idxParentesis = baseTitulo.IndexOf(" (", StringComparison.Ordinal);
        if (idxParentesis > 0)
            baseTitulo = baseTitulo[..idxParentesis].Trim();

        return string.IsNullOrWhiteSpace(baseTitulo) ? "Reservada" : baseTitulo;
    }

    private static bool EsDuracionReservaPublicaValida(TimeOnly horaInicio, TimeOnly horaFin)
    {
        var duracion = (int)(horaFin.ToTimeSpan() - horaInicio.ToTimeSpan()).TotalMinutes;
        if (duracion == DuracionReservaPublicaMinutos)
            return true;

        return horaInicio == HoraInicioCierreMedianoche && horaFin == HoraFinCierreMedianoche;
    }

    private static TimeOnly NormalizarHoraFinReservaPublica(TimeOnly horaInicio, TimeOnly horaFin)
    {
        if (horaInicio != HoraInicioCierreMedianoche)
            return horaFin;

        return HoraFinCierreMedianoche;
    }

    private async Task<ReservaPublicaPageViewModel?> ConstruirReservaVmAsync(
        int espacioDeportivoId,
        DateOnly fecha,
        TimeOnly horaInicio,
        TimeOnly horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId,
        int? negocioId,
        bool omitirFechaHorario = false,
        SolicitudReservaPublicaFormViewModel? formBase = null)
    {
        var codigoDep = string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim();
        var codigoProv = string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim();
        var codigoDist = string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim();

        if (horaFin <= horaInicio)
            horaFin = new TimeOnly(23, 59);

        var disponibles = await spService.HomeBuscarEspaciosDisponiblesAsync(
            fecha,
            horaInicio,
            horaFin,
            codigoDep,
            codigoProv,
            codigoDist,
            tipoDeporteId,
            negocioId,
            omitirFechaHorario);

        var espacio = disponibles.FirstOrDefault(x => x.EspacioDeportivoId == espacioDeportivoId);
        if (espacio is null)
            return null;

        var sedes = await spService.HomeListarSedesAsync();
        var sede = espacio.SedeId.HasValue
            ? sedes.FirstOrDefault(x => x.Id == espacio.SedeId.Value)
            : null;
        var negocioIdResolved = negocioId ?? sede?.NegocioId;
        if (!negocioIdResolved.HasValue)
            return null;

        var tiposDoc = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        var form = formBase ?? new SolicitudReservaPublicaFormViewModel();
        form.EspacioDeportivoId = espacio.EspacioDeportivoId;
        form.Fecha = fecha;
        form.HoraInicio = horaInicio;
        form.HoraFin = horaFin;
        form.CodigoDepartamento = codigoDep;
        form.CodigoProvincia = codigoProv;
        form.CodigoUbigeo = codigoDist;
        form.TipoDeporteId = tipoDeporteId;
        form.NegocioId = negocioIdResolved;
        form.OmitirFechaHorario = omitirFechaHorario;
        form.TipoDocumento = string.IsNullOrWhiteSpace(form.TipoDocumento) ? "0" : form.TipoDocumento.Trim();
        form.TiposDocumentoIdentidad = tiposDoc;

        if (User.Identity?.IsAuthenticated == true)
        {
            var user = await userManager.GetUserAsync(User);
            var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var perfilPublico = !string.IsNullOrWhiteSpace(usuarioId)
                ? await spService.UsuariosPublicosObtenerPerfilAsync(usuarioId)
                : null;

            if (user is not null)
            {
                if (string.IsNullOrWhiteSpace(form.Correo) && !string.IsNullOrWhiteSpace(user.Email))
                    form.Correo = user.Email.Trim();

                if (string.IsNullOrWhiteSpace(form.Nombres) && string.IsNullOrWhiteSpace(form.Apellidos))
                {
                    var nombreRaw = (user.Nombres ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(nombreRaw))
                    {
                        var partes = nombreRaw.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                        if (partes.Length > 1)
                        {
                            form.Nombres = partes[0];
                            form.Apellidos = string.Join(' ', partes.Skip(1));
                        }
                        else
                        {
                        form.Nombres = nombreRaw;
                    }
                }
            }

            if (perfilPublico is not null)
            {
                if (!string.IsNullOrWhiteSpace(perfilPublico.TipoDocumento))
                    form.TipoDocumento = perfilPublico.TipoDocumento;
                if (string.IsNullOrWhiteSpace(form.NumeroDocumento) && !string.IsNullOrWhiteSpace(perfilPublico.NumeroDocumento))
                    form.NumeroDocumento = perfilPublico.NumeroDocumento;
                if (!string.IsNullOrWhiteSpace(perfilPublico.Nombres))
                    form.Nombres = perfilPublico.Nombres;
                if (!string.IsNullOrWhiteSpace(perfilPublico.Apellidos))
                    form.Apellidos = perfilPublico.Apellidos;
                if (string.IsNullOrWhiteSpace(form.NombreEquipo) && !string.IsNullOrWhiteSpace(perfilPublico.NombreEquipo))
                    form.NombreEquipo = perfilPublico.NombreEquipo;
                if (string.IsNullOrWhiteSpace(form.Telefono) && !string.IsNullOrWhiteSpace(perfilPublico.Telefono))
                    form.Telefono = perfilPublico.Telefono;
                if (string.IsNullOrWhiteSpace(form.Correo) && !string.IsNullOrWhiteSpace(perfilPublico.Correo))
                    form.Correo = perfilPublico.Correo;
            }

            form.UsuarioId = usuarioId;
        }
        }

        ReservaCotizacionViewModel? cotizacion = null;
        try
        {
            cotizacion = await spService.ReservasCotizarAsync(
                negocioIdResolved.Value,
                espacio.EspacioDeportivoId,
                fecha,
                horaInicio,
                horaFin);
        }
        catch
        {
            cotizacion = null;
        }

        return new ReservaPublicaPageViewModel
        {
            NegocioId = negocioIdResolved.Value,
            Espacio = espacio,
            Sede = sede,
            Formulario = form,
            Cotizacion = cotizacion
        };
    }

    private async Task<HomeIndexViewModel> ConstruirHomeVmAsync(
        DateOnly? fecha,
        TimeOnly? horaInicio,
        TimeOnly? horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId,
        int? negocioId,
        bool buscarCercaDeMi = false,
        decimal? latitudUsuario = null,
        decimal? longitudUsuario = null,
        decimal? radioKm = null,
        bool omitirFechaHorario = true,
        int pagina = 1)
    {
        const int tamanoPagina = 9;
        var sugerido = ObtenerRangoSugeridoBusqueda(DateTime.Now);
        var fechaConsulta = fecha ?? sugerido.Fecha;
        var horaInicioConsulta = horaInicio ?? sugerido.HoraInicio;
        var horaFinConsulta = horaFin ?? sugerido.HoraFin;
        var codigoDep = string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim();
        var codigoProv = string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim();
        var codigoDist = string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim();
        var omitirHorarioEfectivo = omitirFechaHorario;
        var usarCercania = buscarCercaDeMi && latitudUsuario.HasValue && longitudUsuario.HasValue;
        var radioEfectivo = radioKm is > 0 ? radioKm : 5m;

        if (horaFinConsulta <= horaInicioConsulta)
            horaFinConsulta = new TimeOnly(23, 59);

        var sedes = await spService.HomeListarSedesAsync();
        var deportes = await spService.HomeListarTiposDeporteAsync();
        var banners = await spService.HomeListarBannersPublicosAsync();
        var popupPromociones = await spService.HomeListarPopupPromocionesActivasAsync();
        var paginaSolicitada = Math.Max(1, pagina);
        var espaciosPaginadosResponse = await spService.HomeBuscarEspaciosDisponiblesPaginadoAsync(
            fechaConsulta,
            horaInicioConsulta,
            horaFinConsulta,
            codigoDep,
            codigoProv,
            codigoDist,
            tipoDeporteId,
            negocioId,
            paginaSolicitada,
            tamanoPagina,
            omitirHorarioEfectivo,
            usarCercania,
            latitudUsuario,
            longitudUsuario,
            radioEfectivo);
        var totalResultados = espaciosPaginadosResponse.TotalRegistros;
        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalResultados / (double)tamanoPagina));
        var paginaActual = Math.Clamp(paginaSolicitada, 1, totalPaginas);
        if (paginaActual != paginaSolicitada)
        {
            espaciosPaginadosResponse = await spService.HomeBuscarEspaciosDisponiblesPaginadoAsync(
                fechaConsulta,
                horaInicioConsulta,
                horaFinConsulta,
                codigoDep,
                codigoProv,
                codigoDist,
                tipoDeporteId,
                negocioId,
                paginaActual,
                tamanoPagina,
                omitirHorarioEfectivo,
                usarCercania,
                latitudUsuario,
                longitudUsuario,
                radioEfectivo);
        }
        var espaciosPaginados = PrepararEspaciosParaVista(
            espaciosPaginadosResponse.Espacios,
            sedes,
            fechaConsulta,
            horaInicioConsulta,
            horaFinConsulta,
            negocioId,
            "https://pub-3afaea6b0b354821989565fa4b8bd250.r2.dev/sedes/Sededefecto/complejodefault.webp");
        var departamentos = await spService.UbigeoDepartamentosListarAsync();
        var provincias = !string.IsNullOrWhiteSpace(codigoDep) && codigoDep.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(codigoDep)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
        var distritos = !string.IsNullOrWhiteSpace(codigoProv) && codigoProv.Length == 4
            ? await spService.UbigeoDistritosListarAsync(codigoProv)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
        var negocios = ConstruirNegociosFiltradosPorUbigeo(sedes, codigoDep, codigoProv, codigoDist);
        if (negocioId.HasValue && !negocios.Any(x => x.Value == negocioId.Value.ToString()))
            negocioId = null;
        var portalConfig = await CargarPortalConfigAsync();
        var popupPromocionesConfig = await CargarPopupPromocionesConfigAsync();

        return new HomeIndexViewModel
        {
            Fecha = fechaConsulta,
            HoraInicio = horaInicioConsulta,
            HoraFin = horaFinConsulta,
            CodigoDepartamento = codigoDep,
            CodigoProvincia = codigoProv,
            CodigoUbigeo = codigoDist,
            TipoDeporteId = tipoDeporteId,
            NegocioId = negocioId,
            OmitirFechaHorario = omitirHorarioEfectivo,
            BuscarCercaDeMi = usarCercania,
            LatitudUsuario = latitudUsuario,
            LongitudUsuario = longitudUsuario,
            RadioKm = radioEfectivo,
            DepartamentosUbigeo = departamentos,
            ProvinciasUbigeo = provincias,
            DistritosUbigeo = distritos,
            Negocios = negocios,
            Banners = banners,
            PopupPromociones = popupPromociones,
            Sedes = sedes,
            TiposDeporte = deportes,
            Disponibles = espaciosPaginados,
            PaginaActual = paginaActual,
            TamanoPagina = tamanoPagina,
            TotalResultados = totalResultados,
            TotalPaginas = totalPaginas,
            PortalConfig = portalConfig,
            PopupPromocionesConfig = popupPromocionesConfig
        };
    }

    private static List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem> ConstruirNegociosFiltradosPorUbigeo(
        List<SedePublicaViewModel> sedes,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo)
    {
        var sedesFiltradas = sedes.Where(x =>
            x.NegocioId.HasValue &&
            !string.IsNullOrWhiteSpace(x.NegocioNombre));

        if (!string.IsNullOrWhiteSpace(codigoUbigeo) && codigoUbigeo.Length == 6)
        {
            sedesFiltradas = sedesFiltradas.Where(x =>
                string.Equals(x.CodigoUbigeoNegocio, codigoUbigeo, StringComparison.OrdinalIgnoreCase));
        }
        else if (!string.IsNullOrWhiteSpace(codigoProvincia) && codigoProvincia.Length == 4)
        {
            sedesFiltradas = sedesFiltradas.Where(x =>
                string.Equals(x.CodigoProvinciaNegocio, codigoProvincia, StringComparison.OrdinalIgnoreCase));
        }
        else if (!string.IsNullOrWhiteSpace(codigoDepartamento) && codigoDepartamento.Length == 2)
        {
            sedesFiltradas = sedesFiltradas.Where(x =>
                string.Equals(x.CodigoDepartamentoNegocio, codigoDepartamento, StringComparison.OrdinalIgnoreCase));
        }

        return sedesFiltradas
            .GroupBy(x => new { Id = x.NegocioId!.Value, Nombre = x.NegocioNombre! })
            .OrderBy(x => x.Key.Nombre)
            .Select(x => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem(x.Key.Nombre, x.Key.Id.ToString()))
            .ToList();
    }

    private static List<EspacioDisponibleViewModel> PrepararEspaciosParaVista(
        List<EspacioDisponibleViewModel> espacios,
        List<SedePublicaViewModel> sedes,
        DateOnly fecha,
        TimeOnly horaInicio,
        TimeOnly horaFin,
        int? negocioIdSeleccionado,
        string imagenSedePorDefectoUrl)
    {
        foreach (var item in espacios)
        {
            var sedeRef = item.SedeId.HasValue
                ? sedes.FirstOrDefault(x => x.Id == item.SedeId.Value)
                : null;

            var negocioNombreDestacado = (sedeRef?.NegocioNombre ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(negocioNombreDestacado) && negocioIdSeleccionado.HasValue)
            {
                var sedeNegocio = sedes.FirstOrDefault(x => x.NegocioId == negocioIdSeleccionado.Value);
                negocioNombreDestacado = (sedeNegocio?.NegocioNombre ?? string.Empty).Trim();
            }

            var telefonoContactoTarjeta = (item.TelefonoContacto ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(telefonoContactoTarjeta))
            {
                telefonoContactoTarjeta = (sedeRef?.Telefono ?? string.Empty).Trim();
            }

            var mapaUrl = item.SedeMapaUrl;
            if (string.IsNullOrWhiteSpace(mapaUrl))
            {
                mapaUrl = sedeRef?.GoogleMapsUrl;
            }
            if (string.IsNullOrWhiteSpace(mapaUrl) && sedeRef?.Latitud is decimal lat && sedeRef.Longitud is decimal lng)
            {
                mapaUrl = $"https://www.google.com/maps?q={lat.ToString(CultureInfo.InvariantCulture)},{lng.ToString(CultureInfo.InvariantCulture)}";
            }

            var fotosTarjeta = new List<string>();
            if (!string.IsNullOrWhiteSpace(item.SedeFotoPrincipalUrl))
            {
                fotosTarjeta.Add(item.SedeFotoPrincipalUrl);
            }
            if (item.SedeFotos is not null && item.SedeFotos.Count > 0)
            {
                fotosTarjeta.AddRange(item.SedeFotos.Where(x => !string.IsNullOrWhiteSpace(x)));
            }
            if (!string.IsNullOrWhiteSpace(item.EspacioFotoPrincipalUrl))
            {
                fotosTarjeta.Add(item.EspacioFotoPrincipalUrl);
            }
            if (item.EspacioFotos is not null && item.EspacioFotos.Count > 0)
            {
                fotosTarjeta.AddRange(item.EspacioFotos.Where(x => !string.IsNullOrWhiteSpace(x)));
            }
            fotosTarjeta = fotosTarjeta
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (fotosTarjeta.Count == 0)
            {
                fotosTarjeta.Add(imagenSedePorDefectoUrl);
            }

            var numeroWhatsappEspacio = string.IsNullOrWhiteSpace(item.WhatsappContacto)
                ? string.Empty
                : new string(item.WhatsappContacto.Where(char.IsDigit).ToArray());
            var mensajeWhatsappEspacio = Uri.EscapeDataString(
                $"Hola, quiero reservar {item.NombreEspacio} en la sede {item.SedeNombre} para {fecha:dd/MM/yyyy} de {horaInicio:HH\\:mm} a {horaFin:HH\\:mm}.");
            var enlaceWhatsappEspacio = string.IsNullOrWhiteSpace(numeroWhatsappEspacio)
                ? string.Empty
                : $"https://wa.me/{numeroWhatsappEspacio}?text={mensajeWhatsappEspacio}";

            item.NegocioNombreDestacado = negocioNombreDestacado;
            item.TelefonoContactoResuelto = telefonoContactoTarjeta;
            item.SedeMapaUrlResuelto = mapaUrl;
            item.EnlaceWhatsappEspacio = enlaceWhatsappEspacio;
            item.FotosTarjetaConFallback = fotosTarjeta;
            item.SedeFotosConFallback = fotosTarjeta;
            item.NegocioIdCotizacion = negocioIdSeleccionado ?? sedeRef?.NegocioId;
            item.SedeFacebookUrl = sedeRef?.FacebookUrl;
            item.SedeInstagramUrl = sedeRef?.InstagramUrl;
            item.SedeTwitterUrl = sedeRef?.TwitterUrl;
        }

        return espacios;
    }

    private async Task<PlataformaPortalConfigViewModel> CargarPortalConfigAsync()
    {
        async Task<string?> Get(string key) => await spService.ParametrosGlobalesObtenerValorAsync(key);

        var cfg = new PlataformaPortalConfigViewModel
        {
            BeneficiosTitulo = (await Get("HOME_PORTAL_BENEF_TITULO")) ?? "Todo lo que necesitas para gestionar tus canchas deportivas",
            BeneficiosSubtitulo = (await Get("HOME_PORTAL_BENEF_SUBTITULO")) ?? "SportCenter integra reservas, sedes, pagos y reportes en una sola plataforma para crecer tu operacion.",
            Beneficio1Titulo = (await Get("HOME_PORTAL_BENEF_1_TITULO")) ?? "Sistema de reservas",
            Beneficio1Detalle = (await Get("HOME_PORTAL_BENEF_1_DETALLE")) ?? "Controla la disponibilidad por horario con agenda visual y registro de clientes en segundos.",
            Beneficio2Titulo = (await Get("HOME_PORTAL_BENEF_2_TITULO")) ?? "Multiples sedes",
            Beneficio2Detalle = (await Get("HOME_PORTAL_BENEF_2_DETALLE")) ?? "Administra distintos complejos deportivos desde un solo panel operativo.",
            Beneficio3Titulo = (await Get("HOME_PORTAL_BENEF_3_TITULO")) ?? "Pagos seguros",
            Beneficio3Detalle = (await Get("HOME_PORTAL_BENEF_3_DETALLE")) ?? "Gestiona adelantos, saldos y comprobantes con trazabilidad por reserva.",
            Beneficio4Titulo = (await Get("HOME_PORTAL_BENEF_4_TITULO")) ?? "Promociones especiales",
            Beneficio4Detalle = (await Get("HOME_PORTAL_BENEF_4_DETALLE")) ?? "Crea descuentos por sede, dia y horario para mejorar ocupacion en horas clave.",
            Beneficio5Titulo = (await Get("HOME_PORTAL_BENEF_5_TITULO")) ?? "Estadisticas detalladas",
            Beneficio5Detalle = (await Get("HOME_PORTAL_BENEF_5_DETALLE")) ?? "Analiza ingresos, ocupacion y rendimiento para tomar decisiones con datos.",
            Beneficio6Titulo = (await Get("HOME_PORTAL_BENEF_6_TITULO")) ?? "Mayor visibilidad",
            Beneficio6Detalle = (await Get("HOME_PORTAL_BENEF_6_DETALLE")) ?? "Publica tu negocio en el portal y recibe solicitudes online de nuevos clientes.",
            CtaTitulo = (await Get("HOME_PORTAL_CTA_TITULO")) ?? "Unete a la comunidad de SportCenter",
            CtaSubtitulo = (await Get("HOME_PORTAL_CTA_SUBTITULO")) ?? "Registra tu club deportivo y comienza a gestionar tus canchas de manera eficiente.",
            CtaBotonClubTexto = (await Get("HOME_PORTAL_CTA_BTN_CLUB_TEXTO")) ?? "Registrar mi club",
            CtaBotonClubUrl = (await Get("HOME_PORTAL_CTA_BTN_CLUB_URL")) ?? "/Identity/Account/Register?TipoRegistro=club",
            CtaBotonUsuarioTexto = (await Get("HOME_PORTAL_CTA_BTN_USUARIO_TEXTO")) ?? "Crear cuenta personal",
            CtaBotonUsuarioUrl = (await Get("HOME_PORTAL_CTA_BTN_USUARIO_URL")) ?? "/Identity/Account/Register",
            MarcaTitulo = (await Get("HOME_PORTAL_MARCA_TITULO")) ?? "SportCenter",
            MarcaDescripcion = (await Get("HOME_PORTAL_MARCA_DESC")) ?? "La plataforma lider para la reserva y gestion de canchas deportivas.",
            ContactoEmail = (await Get("HOME_PORTAL_CONTACTO_EMAIL")) ?? "contacto@sportcenter.com",
            ContactoTelefono = (await Get("HOME_PORTAL_CONTACTO_TELEFONO")) ?? "+51 900 000 000",
            SiguenosFacebookUrl = (await Get("HOME_PORTAL_FACEBOOK_URL")) ?? string.Empty,
            SiguenosInstagramUrl = (await Get("HOME_PORTAL_INSTAGRAM_URL")) ?? string.Empty,
            SiguenosWhatsappUrl = (await Get("HOME_PORTAL_WHATSAPP_URL")) ?? string.Empty
        };

        return cfg;
    }

    private async Task<PopupPromocionConfigViewModel> CargarPopupPromocionesConfigAsync()
    {
        async Task<string?> Get(string key) => await spService.ParametrosGlobalesObtenerValorAsync(key);

        return new PopupPromocionConfigViewModel
        {
            ActivarPopupAutomatico = LeerBool(await Get("POPUP_PROMO_AUTO_ENABLED"), true),
            SegundosEsperaAntesDeMostrar = LeerEntero(await Get("POPUP_PROMO_DELAY_SECONDS"), 1, 0, 30),
            ActivarAutoplaySlider = LeerBool(await Get("POPUP_PROMO_AUTOPLAY_ENABLED"), true),
            VelocidadAutoplayMs = LeerEntero(await Get("POPUP_PROMO_AUTOPLAY_MS"), 4500, 1000, 20000),
            MostrarFlechas = LeerBool(await Get("POPUP_PROMO_SHOW_ARROWS"), true),
            MostrarIndicadores = LeerBool(await Get("POPUP_PROMO_SHOW_INDICATORS"), true)
        };
    }

    private static bool LeerBool(string? valor, bool fallback)
    {
        if (string.IsNullOrWhiteSpace(valor))
            return fallback;

        valor = valor.Trim();
        if (valor == "1")
            return true;
        if (valor == "0")
            return false;

        return bool.TryParse(valor, out var parsed) ? parsed : fallback;
    }

    private static int LeerEntero(string? valor, int fallback, int min, int max)
    {
        if (!int.TryParse((valor ?? string.Empty).Trim(), out var parsed))
            return fallback;

        return Math.Clamp(parsed, min, max);
    }
}
