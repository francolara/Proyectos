using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class DesafiosController(
    ISportCenterStoredProcedureService spService,
    IDesafioEmailNotificationService desafioEmailNotificationService) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(int hpage = 1)
    {
        var usuarioId = ObtenerUsuarioId();
        if (usuarioId is null) return Challenge();

        ViewData["PublicFullWidth"] = true;
        var perfil = await spService.UsuariosPublicosObtenerPerfilAsync(usuarioId);
        if (perfil is null || !perfil.BuscarDesafios)
        {
            TempData["PerfilPublicoInfo"] = "Primero activa la opcion 'Buscar desafios' en tu perfil para ingresar al modulo de Desafios.";
            return RedirectToAction("Index", "PerfilPublico", new { tab = "datos" });
        }

        var vm = await ConstruirViewModelAsync(usuarioId, new DesafiosIndexViewModel
        {
            PerfilActual = perfil
        }, incluirBusqueda: false, paginaHistorial: hpage);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BuscarRivales(DesafiosIndexViewModel vm)
    {
        var usuarioId = ObtenerUsuarioId();
        if (usuarioId is null) return Challenge();

        ViewData["PublicFullWidth"] = true;
        LimpiarErroresNuevoDesafio();
        vm = await ConstruirViewModelAsync(usuarioId, vm, incluirBusqueda: true, paginaHistorial: vm.PaginaHistorial);
        return View("Index", vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Crear(
        [Bind(Prefix = "Filtros")] DesafioBusquedaViewModel filtros,
        [Bind(Prefix = "NuevoDesafio")] DesafioCrearViewModel nuevoDesafio)
    {
        var usuarioId = ObtenerUsuarioId();
        if (usuarioId is null) return Challenge();

        ViewData["PublicFullWidth"] = true;
        var vm = await ConstruirViewModelAsync(usuarioId, new DesafiosIndexViewModel
        {
            Filtros = filtros,
            NuevoDesafio = nuevoDesafio
        }, incluirBusqueda: true, paginaHistorial: 1);
        ViewData["AbrirModalDesafio"] = true;

        ValidarNuevoDesafio(vm);
        if (!ModelState.IsValid)
            return View("Index", vm);

        try
        {
            var desafioId = await spService.DesafiosCrearAsync(usuarioId, vm.NuevoDesafio, User.Identity?.Name ?? "desafios");
            if (desafioId <= 0)
            {
                ModelState.AddModelError(string.Empty, "No se pudo registrar el desafio.");
                return View("Index", vm);
            }

            await desafioEmailNotificationService.NotifyDesafioReceivedAsync(desafioId);
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View("Index", vm);
        }

        TempData["DesafiosOk"] = "Desafio enviado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Cancelar(int id)
    {
        var usuarioId = ObtenerUsuarioId();
        if (usuarioId is null) return Challenge();

        await spService.DesafiosCancelarAsync(id, usuarioId, User.Identity?.Name ?? "desafios");
        TempData["DesafiosOk"] = "Desafio cancelado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Responder(int id, string decision)
    {
        var usuarioId = ObtenerUsuarioId();
        if (usuarioId is null) return Challenge();

        var estado = string.Equals(decision, "aceptar", StringComparison.OrdinalIgnoreCase)
            ? DesafioEstados.Aceptado
            : DesafioEstados.Rechazado;

        await spService.DesafiosResponderAsync(id, usuarioId, estado, User.Identity?.Name ?? "desafios");
        TempData["DesafiosOk"] = estado == DesafioEstados.Aceptado
            ? "Desafio aceptado correctamente."
            : "Desafio rechazado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Finalizar(int id)
    {
        var usuarioId = ObtenerUsuarioId();
        if (usuarioId is null) return Challenge();

        await spService.DesafiosFinalizarAsync(id, usuarioId, User.Identity?.Name ?? "desafios");
        TempData["DesafiosOk"] = "Desafio marcado como finalizado.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EnviarMensaje(int id, string mensaje)
    {
        var usuarioId = ObtenerUsuarioId();
        if (usuarioId is null) return Challenge();

        var model = new DesafioMensajeCrearViewModel
        {
            IdDesafio = id,
            Mensaje = string.IsNullOrWhiteSpace(mensaje) ? string.Empty : mensaje.Trim()
        };

        if (model.IdDesafio <= 0 || string.IsNullOrWhiteSpace(model.Mensaje))
        {
            TempData["DesafiosError"] = "Debes escribir un mensaje valido para el desafio.";
            return RedirectToAction(nameof(Index));
        }

        try
        {
            await spService.DesafiosMensajeCrearAsync(usuarioId, model, User.Identity?.Name ?? "desafios");
            TempData["DesafiosOk"] = "Mensaje enviado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["DesafiosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index));
    }

    private async Task<DesafiosIndexViewModel> ConstruirViewModelAsync(string usuarioId, DesafiosIndexViewModel vm, bool incluirBusqueda, int paginaHistorial = 1)
    {
        const int tamanoPaginaHistorial = 4;
        var paginaActualHistorial = paginaHistorial < 1 ? 1 : paginaHistorial;

        vm.PerfilActual = await spService.UsuariosPublicosObtenerPerfilAsync(usuarioId) ?? new UsuarioPublicoPerfilViewModel
        {
            UsuarioId = usuarioId,
            TipoDocumento = "0"
        };

        await CargarCombosAsync(vm);
        vm.DesafiosEnviados = await spService.DesafiosListarAsync(usuarioId, "enviados");
        vm.DesafiosRecibidos = await spService.DesafiosListarAsync(usuarioId, "recibidos");
        var (historial, totalHistorial) = await spService.DesafiosHistorialListarAsync(usuarioId, paginaActualHistorial, tamanoPaginaHistorial);
        var totalPaginasHistorial = Math.Max(1, (int)Math.Ceiling(totalHistorial / (double)tamanoPaginaHistorial));
        if (paginaActualHistorial > totalPaginasHistorial)
        {
            paginaActualHistorial = totalPaginasHistorial;
            (historial, totalHistorial) = await spService.DesafiosHistorialListarAsync(usuarioId, paginaActualHistorial, tamanoPaginaHistorial);
        }

        vm.Historial = historial;
        vm.PaginaHistorial = paginaActualHistorial;
        vm.TamanoPaginaHistorial = tamanoPaginaHistorial;
        vm.TotalHistorial = totalHistorial;
        vm.TotalPaginasHistorial = totalPaginasHistorial;
        AdjuntarMensajes(
            vm,
            await spService.DesafiosMensajesListarAsync(usuarioId));

        if (!incluirBusqueda)
            return vm;

        vm.BusquedaEjecutada = true;

        vm.Filtros.CodigoUbigeo = string.IsNullOrWhiteSpace(vm.Filtros.CodigoUbigeo) ? null : vm.Filtros.CodigoUbigeo.Trim();
        if (string.IsNullOrWhiteSpace(vm.Filtros.CodigoUbigeo))
        {
            ModelState.AddModelError("Filtros.CodigoUbigeo", "Debes seleccionar departamento, provincia y distrito antes de buscar.");
            return vm;
        }

        vm.ResultadosBusqueda = await spService.DesafiosBuscarRivalesAsync(
            usuarioId,
            vm.Filtros.CodigoUbigeo,
            vm.Filtros.IdDeporte,
            vm.Filtros.IdNivel);

        return vm;
    }

    private async Task CargarCombosAsync(DesafiosIndexViewModel vm)
    {
        vm.Filtros.Departamentos = await spService.UbigeoDepartamentosListarAsync();
        vm.Filtros.Provincias = !string.IsNullOrWhiteSpace(vm.Filtros.CodigoDepartamento) && vm.Filtros.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(vm.Filtros.CodigoDepartamento)
            : new();
        vm.Filtros.Distritos = !string.IsNullOrWhiteSpace(vm.Filtros.CodigoProvincia) && vm.Filtros.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(vm.Filtros.CodigoProvincia)
            : new();

        var deportes = (await spService.HomeListarTiposDeporteAsync())
            .Select(x => new SelectListItem(x.Nombre, x.Id.ToString()))
            .ToList();
        var niveles = await spService.DesafiosNivelesListarAsync();

        vm.Filtros.Deportes = deportes;
        vm.Filtros.Niveles = niveles;
        vm.NuevoDesafio.Departamentos = await spService.UbigeoDepartamentosListarAsync();
        vm.NuevoDesafio.Provincias = !string.IsNullOrWhiteSpace(vm.NuevoDesafio.CodigoDepartamento) && vm.NuevoDesafio.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(vm.NuevoDesafio.CodigoDepartamento)
            : new();
        vm.NuevoDesafio.Distritos = !string.IsNullOrWhiteSpace(vm.NuevoDesafio.CodigoProvincia) && vm.NuevoDesafio.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(vm.NuevoDesafio.CodigoProvincia)
            : new();

        if (vm.PerfilActual.DeportesDesafio.Count == 0)
            vm.PerfilActual.DeportesDesafio = deportes;
        if (vm.PerfilActual.NivelesDesafio.Count == 0)
            vm.PerfilActual.NivelesDesafio = niveles;
    }

    private void ValidarNuevoDesafio(DesafiosIndexViewModel vm)
    {
        if (vm.NuevoDesafio.IdUsuarioRetadoPerfil <= 0)
            ModelState.AddModelError("NuevoDesafio.IdUsuarioRetadoPerfil", "Selecciona un rival valido.");

        if (vm.NuevoDesafio.IdUsuarioRetadoPerfil == vm.PerfilActual.Id)
            ModelState.AddModelError("NuevoDesafio.IdUsuarioRetadoPerfil", "No puedes desafiarte a ti mismo.");

        if (vm.NuevoDesafio.FechaTentativa.HasValue && vm.NuevoDesafio.FechaTentativa.Value < DateOnly.FromDateTime(DateTime.Today))
            ModelState.AddModelError("NuevoDesafio.FechaTentativa", "La fecha tentativa no puede ser anterior a hoy.");

        if (vm.NuevoDesafio.HoraTentativa.HasValue &&
            (vm.NuevoDesafio.HoraTentativa.Value.Minute != 0 || vm.NuevoDesafio.HoraTentativa.Value.Second != 0))
            ModelState.AddModelError("NuevoDesafio.HoraTentativa", "La hora tentativa debe seleccionarse en horas exactas.");

        vm.NuevoDesafio.CodigoDepartamento = string.IsNullOrWhiteSpace(vm.NuevoDesafio.CodigoDepartamento)
            ? null
            : vm.NuevoDesafio.CodigoDepartamento.Trim();
        vm.NuevoDesafio.CodigoProvincia = string.IsNullOrWhiteSpace(vm.NuevoDesafio.CodigoProvincia)
            ? null
            : vm.NuevoDesafio.CodigoProvincia.Trim();
        vm.NuevoDesafio.Distrito = string.IsNullOrWhiteSpace(vm.NuevoDesafio.Distrito)
            ? string.Empty
            : vm.NuevoDesafio.Distrito.Trim();

        if (!string.IsNullOrWhiteSpace(vm.NuevoDesafio.Distrito) && vm.NuevoDesafio.Distrito.Length != 6)
            ModelState.AddModelError("NuevoDesafio.Distrito", "Debes seleccionar un distrito valido.");
    }

    private string? ObtenerUsuarioId()
        => User.FindFirstValue(ClaimTypes.NameIdentifier);

    private void LimpiarErroresNuevoDesafio()
    {
        var keys = ModelState.Keys
            .Where(x => x.StartsWith("NuevoDesafio.", StringComparison.Ordinal))
            .ToList();

        foreach (var key in keys)
            ModelState.Remove(key);
    }

    private static void AdjuntarMensajes(DesafiosIndexViewModel vm, List<DesafioMensajeItemViewModel> mensajes)
    {
        var mensajesPorDesafio = mensajes
            .GroupBy(x => x.IdDesafio)
            .ToDictionary(x => x.Key, x => x.OrderBy(y => y.FechaRegistro).ToList());

        AplicarMensajes(vm.DesafiosEnviados, mensajesPorDesafio);
        AplicarMensajes(vm.DesafiosRecibidos, mensajesPorDesafio);
        AplicarMensajes(vm.Historial, mensajesPorDesafio);
    }

    private static void AplicarMensajes(List<DesafioListadoItemViewModel> desafios, IReadOnlyDictionary<int, List<DesafioMensajeItemViewModel>> mensajesPorDesafio)
    {
        foreach (var desafio in desafios)
            desafio.Mensajes = mensajesPorDesafio.TryGetValue(desafio.Id, out var mensajes)
                ? mensajes
                : new List<DesafioMensajeItemViewModel>();
    }
}
