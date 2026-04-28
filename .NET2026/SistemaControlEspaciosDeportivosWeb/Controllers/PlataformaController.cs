using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize(Roles = "OwnerPlataforma")]
public class PlataformaController(
    ISportCenterStoredProcedureService spService,
    IClubRegistrationNotificationService clubRegistrationNotificationService,
    IHomeReferencialesExternosSyncService referencialesExternosSyncService) : Controller
{
    private const string ParamRefExternosBarridoHabilitado = "HOME_REFEXT_BARRIDO_HABILITADO";

    private static readonly (string Key, string Label, string? Desc, Func<PlataformaPortalConfigViewModel, string?> GetValue)[] PortalParamMap =
    [
        ("HOME_PORTAL_BENEF_TITULO", "Beneficios titulo", "Titulo global de la seccion de beneficios del Home.", x => x.BeneficiosTitulo),
        ("HOME_PORTAL_BENEF_SUBTITULO", "Beneficios subtitulo", "Subtitulo global de la seccion de beneficios del Home.", x => x.BeneficiosSubtitulo),
        ("HOME_PORTAL_BENEF_1_TITULO", "Beneficio 1 titulo", null, x => x.Beneficio1Titulo),
        ("HOME_PORTAL_BENEF_1_DETALLE", "Beneficio 1 detalle", null, x => x.Beneficio1Detalle),
        ("HOME_PORTAL_BENEF_2_TITULO", "Beneficio 2 titulo", null, x => x.Beneficio2Titulo),
        ("HOME_PORTAL_BENEF_2_DETALLE", "Beneficio 2 detalle", null, x => x.Beneficio2Detalle),
        ("HOME_PORTAL_BENEF_3_TITULO", "Beneficio 3 titulo", null, x => x.Beneficio3Titulo),
        ("HOME_PORTAL_BENEF_3_DETALLE", "Beneficio 3 detalle", null, x => x.Beneficio3Detalle),
        ("HOME_PORTAL_BENEF_4_TITULO", "Beneficio 4 titulo", null, x => x.Beneficio4Titulo),
        ("HOME_PORTAL_BENEF_4_DETALLE", "Beneficio 4 detalle", null, x => x.Beneficio4Detalle),
        ("HOME_PORTAL_BENEF_5_TITULO", "Beneficio 5 titulo", null, x => x.Beneficio5Titulo),
        ("HOME_PORTAL_BENEF_5_DETALLE", "Beneficio 5 detalle", null, x => x.Beneficio5Detalle),
        ("HOME_PORTAL_BENEF_6_TITULO", "Beneficio 6 titulo", null, x => x.Beneficio6Titulo),
        ("HOME_PORTAL_BENEF_6_DETALLE", "Beneficio 6 detalle", null, x => x.Beneficio6Detalle),
        ("HOME_PORTAL_CTA_TITULO", "CTA titulo", "Titulo del bloque verde del Home.", x => x.CtaTitulo),
        ("HOME_PORTAL_CTA_SUBTITULO", "CTA subtitulo", "Texto de apoyo del bloque verde del Home.", x => x.CtaSubtitulo),
        ("HOME_PORTAL_CTA_BTN_CLUB_TEXTO", "CTA boton club texto", null, x => x.CtaBotonClubTexto),
        ("HOME_PORTAL_CTA_BTN_CLUB_URL", "CTA boton club URL", null, x => x.CtaBotonClubUrl),
        ("HOME_PORTAL_CTA_BTN_USUARIO_TEXTO", "CTA boton usuario texto", null, x => x.CtaBotonUsuarioTexto),
        ("HOME_PORTAL_CTA_BTN_USUARIO_URL", "CTA boton usuario URL", null, x => x.CtaBotonUsuarioUrl),
        ("HOME_PORTAL_MARCA_TITULO", "Marca titulo", "Nombre de plataforma en pie del Home.", x => x.MarcaTitulo),
        ("HOME_PORTAL_MARCA_DESC", "Marca descripcion", null, x => x.MarcaDescripcion),
        ("HOME_PORTAL_CONTACTO_EMAIL", "Contacto email", null, x => x.ContactoEmail),
        ("HOME_PORTAL_CONTACTO_TELEFONO", "Contacto telefono", null, x => x.ContactoTelefono),
        ("HOME_PORTAL_FACEBOOK_URL", "Facebook URL", null, x => x.SiguenosFacebookUrl),
        ("HOME_PORTAL_INSTAGRAM_URL", "Instagram URL", null, x => x.SiguenosInstagramUrl),
        ("HOME_PORTAL_WHATSAPP_URL", "WhatsApp URL", null, x => x.SiguenosWhatsappUrl),
        ("HOME_PORTAL_NOTIF_CORREO_1", "Notificaciones correo 1", "Correo principal para futuras notificaciones internas del portal web.", x => x.NotificacionCorreo1),
        ("HOME_PORTAL_NOTIF_CORREO_2", "Notificaciones correo 2", "Correo secundario para futuras notificaciones internas del portal web.", x => x.NotificacionCorreo2)
    ];

    public async Task<IActionResult> Index()
    {
        ViewData["PlatformShell"] = true;

        var banners = await spService.BannersAdminListarAsync(null);
        var vm = new PlataformaIndexViewModel
        {
            CorreoUsuario = User.Identity?.Name ?? string.Empty,
            TotalBanners = banners.Count,
            BannersActivos = banners.Count(x => x.Activo),
            BannersInactivos = banners.Count(x => !x.Activo)
        };

        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> PortalWeb()
    {
        ViewData["PlatformShell"] = true;
        var vm = await CargarPortalConfigAsync();
        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> Negocios(string? buscar = null, string? estadoContrato = null, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        const int tamanoPagina = 20;
        var estadoContratoNormalizado = NormalizarEstadoContrato(estadoContrato);
        var paginaActual = pagina < 1 ? 1 : pagina;
        var resultado = await spService.PlataformaNegociosListarAsync(buscar, estadoContratoNormalizado, paginaActual, tamanoPagina);
        var totalRegistros = resultado.TotalRegistros;
        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)tamanoPagina));
        if (paginaActual > totalPaginas)
            paginaActual = totalPaginas;
        if (paginaActual != pagina)
            resultado = await spService.PlataformaNegociosListarAsync(buscar, estadoContratoNormalizado, paginaActual, tamanoPagina);

        var vm = new PlataformaNegociosAdminViewModel
        {
            Buscar = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(),
            EstadoContrato = estadoContratoNormalizado,
            Pagina = paginaActual,
            TamanoPagina = tamanoPagina,
            TotalRegistros = totalRegistros,
            TotalPaginas = totalPaginas,
            Negocios = resultado.Negocios
        };
        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> ClubesPendientes(int? estado = 1, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        const int tamanoPagina = 20;
        var paginaActual = pagina < 1 ? 1 : pagina;
        var resultado = await spService.AltasClubesListarAsync(estado, paginaActual, tamanoPagina);
        var totalRegistros = resultado.TotalRegistros;
        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)tamanoPagina));
        paginaActual = Math.Clamp(paginaActual, 1, totalPaginas);
        if (paginaActual != pagina)
            resultado = await spService.AltasClubesListarAsync(estado, paginaActual, tamanoPagina);
        var vm = new PlataformaAltasClubesAdminViewModel
        {
            Estado = estado,
            DiasPruebaDefault = 30,
            Pagina = paginaActual,
            TamanoPagina = tamanoPagina,
            TotalRegistros = totalRegistros,
            TotalPaginas = totalPaginas,
            TotalPendientes = resultado.TotalPendientes,
            TotalAprobados = resultado.TotalAprobados,
            TotalRechazados = resultado.TotalRechazados,
            Solicitudes = resultado.Solicitudes
        };
        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> ReferencialesExternos(
        string? buscarNombre = null,
        string? filtroCodigoDepartamento = null,
        string? filtroCodigoProvincia = null,
        string? filtroCodigoUbigeo = null,
        bool incluirInactivos = false,
        int paginaListado = 1)
    {
        ViewData["PlatformShell"] = true;
        var vm = new PlataformaReferencialesExternosViewModel
        {
            BuscarNombre = string.IsNullOrWhiteSpace(buscarNombre) ? null : buscarNombre.Trim(),
            FiltroCodigoDepartamento = string.IsNullOrWhiteSpace(filtroCodigoDepartamento) ? null : filtroCodigoDepartamento.Trim(),
            FiltroCodigoProvincia = string.IsNullOrWhiteSpace(filtroCodigoProvincia) ? null : filtroCodigoProvincia.Trim(),
            FiltroCodigoUbigeo = string.IsNullOrWhiteSpace(filtroCodigoUbigeo) ? null : filtroCodigoUbigeo.Trim(),
            IncluirInactivos = incluirInactivos,
            PaginaListado = paginaListado <= 0 ? 1 : paginaListado,
            TamanoPaginaListado = 50
        };
        vm = await CargarReferencialesExternosVmAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ReferencialesExternos(PlataformaReferencialesExternosViewModel model)
    {
        ViewData["PlatformShell"] = true;
        if (!await ObtenerFlagBarridoReferencialesAsync())
        {
            TempData["PortalWebError"] = "El barrido manual de referenciales externos esta deshabilitado por configuracion global.";
            var vmDisabled = await CargarReferencialesExternosVmAsync(model);
            return View(vmDisabled);
        }

        model.PalabraClave = (model.PalabraClave ?? string.Empty).Trim();
        model.CodigoDepartamento = (model.CodigoDepartamento ?? string.Empty).Trim();
        model.CodigoProvincia = (model.CodigoProvincia ?? string.Empty).Trim();
        model.CodigoUbigeo = (model.CodigoUbigeo ?? string.Empty).Trim();
        model.MaxResultados = Math.Clamp(model.MaxResultados, 1, 60);

        if (model.TipoDeporteSuperId <= 0)
            ModelState.AddModelError(nameof(model.TipoDeporteSuperId), "Debes seleccionar un tipo de deporte valido.");
        if (string.IsNullOrWhiteSpace(model.PalabraClave))
            ModelState.AddModelError(nameof(model.PalabraClave), "Debes ingresar una palabra clave.");

        if (model.CodigoDepartamento.Length != 2)
            ModelState.AddModelError(nameof(model.CodigoDepartamento), "Debes seleccionar un departamento valido.");
        if (model.CodigoProvincia.Length != 4)
            ModelState.AddModelError(nameof(model.CodigoProvincia), "Debes seleccionar una provincia valida.");
        if (model.CodigoUbigeo.Length != 6)
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "Debes seleccionar un distrito valido.");

        if (!ModelState.IsValid)
        {
            var vmInvalid = await CargarReferencialesExternosVmAsync(model);
            return View(vmInvalid);
        }

        try
        {
            model.Resultado = await referencialesExternosSyncService.EjecutarBarridoAsync(
                model.CodigoUbigeo,
                model.TipoDeporteSuperId,
                model.PalabraClave,
                model.MaxResultados,
                model.DescargarTelefonos,
                model.DescargarFotos,
                User.Identity?.Name ?? "owner-platform");

            TempData["PortalWebOk"] = "Barrido ejecutado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["PortalWebError"] = $"No se pudo ejecutar el barrido: {ex.Message}";
        }

        var vm = await CargarReferencialesExternosVmAsync(model);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ReferencialesExternosInactivar(
        int id,
        string? buscarNombre = null,
        string? filtroCodigoDepartamento = null,
        string? filtroCodigoProvincia = null,
        string? filtroCodigoUbigeo = null,
        bool incluirInactivos = false,
        int paginaListado = 1)
    {
        ViewData["PlatformShell"] = true;
        var ok = await spService.HomeReferencialesExternosInactivarAsync(id, User.Identity?.Name ?? "owner-platform");
        TempData[ok ? "PortalWebOk" : "PortalWebError"] = ok
            ? "Referencial externo inactivado."
            : "No se pudo inactivar el referencial seleccionado.";

        return RedirectToAction(
            nameof(ReferencialesExternos),
            new
            {
                buscarNombre,
                filtroCodigoDepartamento,
                filtroCodigoProvincia,
                filtroCodigoUbigeo,
                incluirInactivos,
                paginaListado
            });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarLimitesNegocio(int negocioId, int sedesPermitidas, int espaciosPermitidos, int usuariosPermitidos, string? buscar = null, string? estadoContrato = null, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        sedesPermitidas = Math.Max(1, sedesPermitidas);
        espaciosPermitidos = Math.Max(1, espaciosPermitidos);
        usuariosPermitidos = Math.Max(1, usuariosPermitidos);
        var ok = await spService.PlataformaNegocioActualizarLimitesAsync(
            negocioId,
            sedesPermitidas,
            espaciosPermitidos,
            usuariosPermitidos,
            User.Identity?.Name ?? "owner-platform");

        TempData[ok ? "PortalWebOk" : "PortalWebError"] = ok
            ? "Limites del negocio actualizados."
            : "No se pudo actualizar los limites del negocio.";

        return RedirectToAction(nameof(Negocios), new { buscar, estadoContrato, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ActivarContratoNegocio(int negocioId, string tipoCobro, DateOnly fechaDesde, DateOnly fechaHasta, int diasGracia = 5, string? buscar = null, string? estadoContrato = null, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        fechaHasta = CalcularFechaFinContrato(tipoCobro, fechaDesde);
        if (fechaHasta < fechaDesde)
        {
            TempData["PortalWebError"] = "La fecha fin no puede ser menor a la fecha inicio.";
            return RedirectToAction(nameof(Negocios), new { buscar, estadoContrato, pagina });
        }

        var ok = await spService.PlataformaNegocioActivarContratoAsync(
            negocioId,
            tipoCobro,
            fechaDesde,
            fechaHasta,
            diasGracia <= 0 ? 5 : diasGracia,
            User.Identity?.Name ?? "owner-platform");

        TempData[ok ? "PortalWebOk" : "PortalWebError"] = ok
            ? "Contrato activado correctamente."
            : "No se pudo activar el contrato del negocio.";

        return RedirectToAction(nameof(Negocios), new { buscar, estadoContrato, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> RenovarContratoNegocio(int negocioId, string? buscar = null, string? estadoContrato = null, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        var ok = await spService.PlataformaNegocioRenovarContratoAsync(
            negocioId,
            User.Identity?.Name ?? "owner-platform");

        TempData[ok ? "PortalWebOk" : "PortalWebError"] = ok
            ? "Suscripcion renovada correctamente."
            : "No se pudo renovar la suscripcion del negocio.";

        return RedirectToAction(nameof(Negocios), new { buscar, estadoContrato, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> FinalizarContratoNegocio(int negocioId, string? buscar = null, string? estadoContrato = null, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        var ok = await spService.PlataformaNegocioFinalizarContratoAsync(
            negocioId,
            User.Identity?.Name ?? "owner-platform");

        TempData[ok ? "PortalWebOk" : "PortalWebError"] = ok
            ? "Contrato finalizado correctamente."
            : "No se pudo finalizar el contrato del negocio.";

        return RedirectToAction(nameof(Negocios), new { buscar, estadoContrato, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AprobarClub(int id, int diasPrueba = 30, string? comentarioGestion = null, int? estado = 1, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        var diasPruebaNormalizado = diasPrueba <= 0 ? 30 : diasPrueba;
        var solicitud = await spService.AltasClubesObtenerPorIdAsync(id);
        if (solicitud is null)
        {
            TempData["PortalWebError"] = "No se encontro la solicitud seleccionada.";
            return RedirectToAction(nameof(ClubesPendientes), new { estado, pagina });
        }

        var ok = await spService.AltasClubesAprobarAsync(
            id,
            User.Identity?.Name ?? "owner-platform",
            comentarioGestion,
            diasPruebaNormalizado);

        if (ok)
        {
            await clubRegistrationNotificationService.NotifyClubApprovalAsync(solicitud, diasPruebaNormalizado);
        }

        TempData[ok ? "PortalWebOk" : "PortalWebError"] = ok
            ? "Solicitud aprobada y negocio activado con periodo de prueba."
            : "No se pudo aprobar la solicitud.";
        return RedirectToAction(nameof(ClubesPendientes), new { estado, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> RechazarClub(int id, string? comentarioGestion = null, int? estado = 1, int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        var ok = await spService.AltasClubesRechazarAsync(
            id,
            User.Identity?.Name ?? "owner-platform",
            comentarioGestion);

        TempData[ok ? "PortalWebOk" : "PortalWebError"] = ok
            ? "Solicitud rechazada."
            : "No se pudo rechazar la solicitud.";
        return RedirectToAction(nameof(ClubesPendientes), new { estado, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> PortalWebGuardar(PlataformaPortalConfigViewModel model)
    {
        ViewData["PlatformShell"] = true;

        if (!ModelState.IsValid)
            return View(model);

        try
        {
            var usuario = User.Identity?.Name ?? "owner-platform";
            foreach (var p in PortalParamMap)
            {
                var esperado = (p.GetValue(model) ?? string.Empty).Trim();
                await spService.ParametrosGlobalesUpsertValorAsync(p.Key, p.Desc, esperado, usuario);
                var actual = (await spService.ParametrosGlobalesObtenerValorAsync(p.Key) ?? string.Empty).Trim();
                if (!string.Equals(actual, esperado, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException($"No se pudo confirmar la persistencia del parametro {p.Key}.");
                }
            }

            TempData["PortalWebOk"] = "Configuracion del portal actualizada.";
            return RedirectToAction(nameof(PortalWeb));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, $"No se pudo guardar la configuracion: {ex.Message}");
            return View(model);
        }
    }

    private async Task<PlataformaPortalConfigViewModel> CargarPortalConfigAsync()
    {
        async Task<string?> Get(string key) => await spService.ParametrosGlobalesObtenerValorAsync(key);

        return new PlataformaPortalConfigViewModel
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
            SiguenosWhatsappUrl = (await Get("HOME_PORTAL_WHATSAPP_URL")) ?? string.Empty,
            NotificacionCorreo1 = (await Get("HOME_PORTAL_NOTIF_CORREO_1")) ?? string.Empty,
            NotificacionCorreo2 = (await Get("HOME_PORTAL_NOTIF_CORREO_2")) ?? string.Empty
        };
    }

    private static string NormalizarEstadoContrato(string? estadoContrato)
    {
        var valor = (estadoContrato ?? "todos").Trim().ToLowerInvariant();
        return valor switch
        {
            "con-contrato" => "con-contrato",
            "sin-contrato" => "sin-contrato",
            "prueba-por-vencer" => "prueba-por-vencer",
            _ => "todos"
        };
    }

    private static DateOnly CalcularFechaFinContrato(string? tipoCobro, DateOnly fechaDesde)
    {
        var tipoCobroNormalizado = (tipoCobro ?? "MENSUAL").Trim().ToUpperInvariant();
        return tipoCobroNormalizado switch
        {
            "TRIMESTRAL" => fechaDesde.AddMonths(3),
            "SEMESTRAL" => fechaDesde.AddMonths(6),
            "ANUAL" => fechaDesde.AddYears(1),
            _ => fechaDesde.AddMonths(1)
        };
    }

    private async Task<PlataformaReferencialesExternosViewModel> CargarReferencialesExternosVmAsync(PlataformaReferencialesExternosViewModel model)
    {
        var codigoDep = (model.CodigoDepartamento ?? string.Empty).Trim();
        var codigoProv = (model.CodigoProvincia ?? string.Empty).Trim();
        var filtroCodigoDep = (model.FiltroCodigoDepartamento ?? string.Empty).Trim();
        var filtroCodigoProv = (model.FiltroCodigoProvincia ?? string.Empty).Trim();
        model.BuscarNombre = string.IsNullOrWhiteSpace(model.BuscarNombre) ? null : model.BuscarNombre.Trim();
        model.PaginaListado = model.PaginaListado <= 0 ? 1 : model.PaginaListado;
        model.TamanoPaginaListado = model.TamanoPaginaListado <= 0 ? 50 : model.TamanoPaginaListado;

        model.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();
        model.ProvinciasUbigeo = codigoDep.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(codigoDep)
            : new List<SelectListItem>();
        model.DistritosUbigeo = codigoProv.Length == 4
            ? await spService.UbigeoDistritosListarAsync(codigoProv)
            : new List<SelectListItem>();

        model.TiposDeporte = await spService.HomeReferencialesExternosListarTiposDeporteSuperAsync();
        model.FiltroDepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();
        model.FiltroProvinciasUbigeo = filtroCodigoDep.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(filtroCodigoDep)
            : new List<SelectListItem>();
        model.FiltroDistritosUbigeo = filtroCodigoProv.Length == 4
            ? await spService.UbigeoDistritosListarAsync(filtroCodigoProv)
            : new List<SelectListItem>();

        var (items, totalRegistros) = await spService.HomeReferencialesExternosListarAdminAsync(
            model.FiltroCodigoDepartamento,
            model.FiltroCodigoProvincia,
            model.FiltroCodigoUbigeo,
            model.BuscarNombre,
            model.PaginaListado,
            model.TamanoPaginaListado,
            model.IncluirInactivos ? null : true);

        model.ReferencialesListado = items;
        model.TotalRegistrosListado = totalRegistros;
        model.TotalPaginasListado = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)model.TamanoPaginaListado));
        if (model.PaginaListado > model.TotalPaginasListado)
            model.PaginaListado = model.TotalPaginasListado;

        if (model.ReferencialesListado.Count == 0 && totalRegistros > 0)
        {
            var (itemsRecalc, _) = await spService.HomeReferencialesExternosListarAdminAsync(
                model.FiltroCodigoDepartamento,
                model.FiltroCodigoProvincia,
                model.FiltroCodigoUbigeo,
                model.BuscarNombre,
                model.PaginaListado,
                model.TamanoPaginaListado,
                model.IncluirInactivos ? null : true);
            model.ReferencialesListado = itemsRecalc;
        }

        model.BarridoHabilitado = await ObtenerFlagBarridoReferencialesAsync();

        return model;
    }

    private async Task<bool> ObtenerFlagBarridoReferencialesAsync()
    {
        var raw = (await spService.ParametrosGlobalesObtenerValorAsync(ParamRefExternosBarridoHabilitado) ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(raw)) return false;

        return raw.Equals("1", StringComparison.OrdinalIgnoreCase)
               || raw.Equals("true", StringComparison.OrdinalIgnoreCase)
               || raw.Equals("si", StringComparison.OrdinalIgnoreCase)
               || raw.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }
}
