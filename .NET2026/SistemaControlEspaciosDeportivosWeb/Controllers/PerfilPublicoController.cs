using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class PerfilPublicoController(ISportCenterStoredProcedureService spService) : Controller
{
    [HttpGet]
    public IActionResult FaqUsuario()
    {
        ViewData["PublicFullWidth"] = true;
        return View();
    }

    [HttpGet]
    public async Task<IActionResult> Index(string? tab = null, int pagina = 1)
    {
        ViewData["PublicFullWidth"] = true;
        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId)) return Challenge();
        const int tamanoPaginaReservas = 6;
        var paginaActualReservas = pagina < 1 ? 1 : pagina;

        var perfil = await spService.UsuariosPublicosObtenerPerfilAsync(usuarioId) ?? new UsuarioPublicoPerfilViewModel
        {
            UsuarioId = usuarioId,
            TipoDocumento = "0"
        };

        InicializarTelefonosParaVista(perfil);
        await CargarCombosAsync(perfil);
        var (reservas, totalReservas) = await spService.UsuariosPublicosReservasListarAsync(usuarioId, paginaActualReservas, tamanoPaginaReservas);
        var aliasSugerido = ConstruirAliasPublicoSugerido(perfil.Nombres, perfil.Apellidos);
        foreach (var reserva in reservas)
        {
            reserva.AliasResenaSugerido = aliasSugerido;
        }

        var totalPaginasReservas = Math.Max(1, (int)Math.Ceiling(totalReservas / (double)tamanoPaginaReservas));
        if (paginaActualReservas > totalPaginasReservas)
        {
            paginaActualReservas = totalPaginasReservas;
            (reservas, totalReservas) = await spService.UsuariosPublicosReservasListarAsync(usuarioId, paginaActualReservas, tamanoPaginaReservas);
            foreach (var reserva in reservas)
            {
                reserva.AliasResenaSugerido = aliasSugerido;
            }
        }

        ViewData["Tab"] = string.IsNullOrWhiteSpace(tab) ? "datos" : tab.Trim().ToLowerInvariant();
        return View(new PerfilPublicoIndexViewModel
        {
            Perfil = perfil,
            Reservas = reservas,
            PaginaReservas = paginaActualReservas,
            TamanoPaginaReservas = tamanoPaginaReservas,
            TotalReservas = totalReservas,
            TotalPaginasReservas = totalPaginasReservas
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarResena(UsuarioPublicoResenaGuardarViewModel model)
    {
        ViewData["PublicFullWidth"] = true;
        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId)) return Challenge();

        model.Pagina = model.Pagina < 1 ? 1 : model.Pagina;
        model.AliasPublico = (model.AliasPublico ?? string.Empty).Trim();
        model.Comentario = (model.Comentario ?? string.Empty).Trim();

        var perfil = await spService.UsuariosPublicosObtenerPerfilAsync(usuarioId);
        var aliasSugerido = ConstruirAliasPublicoSugerido(perfil?.Nombres, perfil?.Apellidos);
        if (string.IsNullOrWhiteSpace(model.AliasPublico))
        {
            model.AliasPublico = aliasSugerido;
        }

        if (!ModelState.IsValid)
        {
            TempData["PerfilPublicoInfo"] = "No se pudo registrar la resena. Revisa el alias visible y el comentario.";
            return RedirectToAction(nameof(Index), new { tab = "reservas", pagina = model.Pagina });
        }

        try
        {
            await spService.UsuariosPublicosResenaCrearAsync(usuarioId, model, User.Identity?.Name ?? "perfil-publico");
            TempData["PerfilPublicoOk"] = "Resena registrada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["PerfilPublicoInfo"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { tab = "reservas", pagina = model.Pagina });
    }

    [HttpGet]
    public async Task<IActionResult> DescargarCalendarioReserva(int reservaId)
    {
        ViewData["PublicFullWidth"] = true;
        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId)) return Challenge();
        if (reservaId <= 0) return NotFound();

        var reserva = await spService.UsuariosPublicosReservaCalendarioObtenerAsync(usuarioId, reservaId);
        if (reserva is null) return NotFound();
        if (reserva.EstadoId is 5 or 6)
        {
            TempData["PerfilPublicoInfo"] = "Solo puedes agregar al calendario reservas activas.";
            return RedirectToAction(nameof(Index), new { tab = "reservas", pagina = 1 });
        }

        var bytes = ReservationCalendarIcsBuilder.Build(reserva, DateTime.UtcNow);
        return File(bytes, "text/calendar; charset=utf-8", $"reserva-{reserva.CodigoReserva}.ics");
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarDatos(PerfilPublicoIndexViewModel vm)
    {
        ViewData["PublicFullWidth"] = true;
        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId)) return Challenge();

        vm.Perfil.UsuarioId = usuarioId;
        vm.Perfil.TipoDocumento = string.IsNullOrWhiteSpace(vm.Perfil.TipoDocumento) ? "0" : vm.Perfil.TipoDocumento.Trim();
        vm.Perfil.NumeroDocumento = string.IsNullOrWhiteSpace(vm.Perfil.NumeroDocumento) ? null : vm.Perfil.NumeroDocumento.Trim();
        vm.Perfil.Nombres = (vm.Perfil.Nombres ?? string.Empty).Trim();
        vm.Perfil.Apellidos = (vm.Perfil.Apellidos ?? string.Empty).Trim();
        vm.Perfil.NombreEquipo = string.IsNullOrWhiteSpace(vm.Perfil.NombreEquipo) ? null : vm.Perfil.NombreEquipo.Trim();
        vm.Perfil.Correo = string.IsNullOrWhiteSpace(vm.Perfil.Correo) ? null : vm.Perfil.Correo.Trim();
        vm.Perfil.CodigoUbigeo = string.IsNullOrWhiteSpace(vm.Perfil.CodigoUbigeo) ? null : vm.Perfil.CodigoUbigeo.Trim();
        vm.Perfil.ObservacionDesafio = string.IsNullOrWhiteSpace(vm.Perfil.ObservacionDesafio) ? null : vm.Perfil.ObservacionDesafio.Trim();
        vm.Perfil.DetalleEquipo = string.IsNullOrWhiteSpace(vm.Perfil.DetalleEquipo) ? null : vm.Perfil.DetalleEquipo.Trim();
        vm.Perfil.CodigoUbigeoEquipo = string.IsNullOrWhiteSpace(vm.Perfil.CodigoUbigeoEquipo) ? null : vm.Perfil.CodigoUbigeoEquipo.Trim();
        ComponerTelefonos(vm.Perfil);

        if (vm.Perfil.BuscarDesafios)
        {
            if (string.IsNullOrWhiteSpace(vm.Perfil.CodigoUbigeoEquipo))
                ModelState.AddModelError("Perfil.CodigoUbigeoEquipo", "Debes seleccionar la ubicacion del equipo para habilitar los desafios.");
            if (!vm.Perfil.IdDeporteDesafio.HasValue || vm.Perfil.IdDeporteDesafio <= 0)
                ModelState.AddModelError("Perfil.IdDeporteDesafio", "Debes seleccionar un deporte para los desafios.");
            if (!vm.Perfil.IdNivelDesafio.HasValue || vm.Perfil.IdNivelDesafio <= 0)
                ModelState.AddModelError("Perfil.IdNivelDesafio", "Debes seleccionar un nivel para los desafios.");
        }

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(vm.Perfil);
            var (reservas, totalReservas) = await spService.UsuariosPublicosReservasListarAsync(usuarioId, 1, 6);
            vm.Reservas = reservas;
            vm.PaginaReservas = 1;
            vm.TamanoPaginaReservas = 6;
            vm.TotalReservas = totalReservas;
            vm.TotalPaginasReservas = Math.Max(1, (int)Math.Ceiling(totalReservas / 6d));
            ViewData["Tab"] = "datos";
            return View("Index", vm);
        }

        await spService.UsuariosPublicosGuardarPerfilAsync(vm.Perfil, User.Identity?.Name ?? "perfil-publico");
        TempData["PerfilPublicoOk"] = "Perfil publico actualizado correctamente.";
        return RedirectToAction(nameof(Index), new { tab = "datos" });
    }

    private async Task CargarCombosAsync(UsuarioPublicoPerfilViewModel perfil)
    {
        perfil.TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        if (perfil.TiposDocumento.All(x => x.Value != "0"))
            perfil.TiposDocumento.Insert(0, new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem("Doc. trib. no dom. sin RUC (0)", "0"));

        perfil.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();
        perfil.ProvinciasUbigeo = !string.IsNullOrWhiteSpace(perfil.CodigoDepartamento) && perfil.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(perfil.CodigoDepartamento)
            : new();
        perfil.DistritosUbigeo = !string.IsNullOrWhiteSpace(perfil.CodigoProvincia) && perfil.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(perfil.CodigoProvincia)
            : new();
        perfil.DepartamentosUbigeoEquipo = await spService.UbigeoDepartamentosListarAsync();
        perfil.ProvinciasUbigeoEquipo = !string.IsNullOrWhiteSpace(perfil.CodigoDepartamentoEquipo) && perfil.CodigoDepartamentoEquipo.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(perfil.CodigoDepartamentoEquipo)
            : new();
        perfil.DistritosUbigeoEquipo = !string.IsNullOrWhiteSpace(perfil.CodigoProvinciaEquipo) && perfil.CodigoProvinciaEquipo.Length == 4
            ? await spService.UbigeoDistritosListarAsync(perfil.CodigoProvinciaEquipo)
            : new();
        perfil.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(perfil.TelefonoCodigoPais);
        perfil.DeportesDesafio = (await spService.HomeListarTiposDeporteAsync())
            .Select(x => new SelectListItem(x.Nombre, x.Id.ToString()))
            .ToList();
        perfil.NivelesDesafio = await spService.DesafiosNivelesListarAsync();
    }

    private static void InicializarTelefonosParaVista(UsuarioPublicoPerfilViewModel perfil)
    {
        TelefonoInternacionalHelper.Descomponer(perfil.Telefono, out var telefonoCodigoPais, out var telefonoNumeroLocal);
        perfil.TelefonoCodigoPais = telefonoCodigoPais;
        perfil.TelefonoNumeroLocal = telefonoNumeroLocal;

        TelefonoInternacionalHelper.Descomponer(perfil.WhatsappEquipo, out var whatsappCodigoPais, out var whatsappNumeroLocal);
        perfil.WhatsappCodigoPais = whatsappCodigoPais;
        perfil.WhatsappNumeroLocal = whatsappNumeroLocal;
    }

    private static void ComponerTelefonos(UsuarioPublicoPerfilViewModel perfil)
    {
        perfil.Telefono = TelefonoInternacionalHelper.Componer(perfil.TelefonoCodigoPais, perfil.TelefonoNumeroLocal);
        perfil.WhatsappEquipo = TelefonoInternacionalHelper.Componer(perfil.WhatsappCodigoPais, perfil.WhatsappNumeroLocal);
    }

    private static string ConstruirAliasPublicoSugerido(string? nombres, string? apellidos)
    {
        static string ObtenerPrimerToken(string? valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return string.Empty;

            return valor
                .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .FirstOrDefault() ?? string.Empty;
        }

        var primerNombre = ObtenerPrimerToken(nombres);
        var primerApellido = ObtenerPrimerToken(apellidos);
        var aliasBase = string.Concat(primerNombre, primerApellido).Replace(" ", string.Empty).Trim();

        return string.IsNullOrWhiteSpace(aliasBase) ? "@JugadorAnonimo" : $"@{aliasBase}";
    }
}
