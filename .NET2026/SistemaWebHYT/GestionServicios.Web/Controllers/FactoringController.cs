using System.Security.Claims;
using System.Text.Json;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Billing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 10/11/2026 | Listado, registro e impresión de negociaciones de factoring.
[Authorize(Policy = "perm:FACTORING.CONSULTAR")]
[Route("Administracion/Factoring")]
public sealed class FactoringController(IFactoringRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private DateOnly Today => DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
        TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "FACTORING.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(DateOnly? desde, DateOnly? hasta, string? buscar, int pagina = 1, CancellationToken ct = default)
    {
        var filter = Filter(desde, hasta, buscar, pagina);
        if (!ModelState.IsValid || !Valid(filter)) return BadRequest("Revisa los filtros.");
        return View(new FactoringList(filter, await ReadAsync(filter, false, ct), CanEdit));
    }

    [HttpGet("Nuevo")]
    [Authorize(Policy = "perm:FACTORING.EDITAR")]
    public async Task<IActionResult> New(DateOnly? desde, DateOnly? hasta, string? buscar, CancellationToken ct)
    {
        var filter = Filter(desde, hasta, buscar);
        if (!ModelState.IsValid || !Valid(filter)) return BadRequest("Revisa los filtros.");
        return View(new FactoringForm(filter, await ReadAsync(filter, true, ct), new FactoringInput { FechaNegociacion = Today }));
    }

    [HttpPost("Guardar")]
    [Authorize(Policy = "perm:FACTORING.EDITAR")]
    public async Task<IActionResult> Save(FactoringInput input, DateOnly? desde, DateOnly? hasta, string? buscar, CancellationToken ct)
    {
        var filter = Filter(desde, hasta, buscar);
        if (!Valid(filter)) return BadRequest("Revisa los filtros.");
        if (input.FechaNegociacion == default || input.FechaDesembolso < input.FechaNegociacion)
            ModelState.AddModelError("", "Revisa las fechas de negociación y desembolso.");
        if (input.Facturas.Count is < 1 or > 100 || input.Facturas.Any(id => id <= 0))
            ModelState.AddModelError("", "Selecciona entre 1 y 100 facturas.");
        if (input.OperacionId == Guid.Empty || decimal.Round(input.Porcentaje, 4) != input.Porcentaje)
            ModelState.AddModelError("", "Revisa la operación y el porcentaje. Se admiten hasta cuatro decimales.");
        if (ModelState.IsValid)
        {
            try
            {
                var id = await repository.SaveFactoringAsync(Actor, input.ToLot(), ct);
                return RedirectToAction(nameof(Details), new { id });
            }
            catch (MaintenanceException error) { ModelState.AddModelError("", error.Message); }
        }
        return View("New", new FactoringForm(filter, await ReadAsync(filter, true, ct), input));
    }

    [HttpGet("{id:long}")]
    public async Task<IActionResult> Details(long id, bool print, CancellationToken ct)
    {
        if (id <= 0) return NotFound();
        var data = await repository.ReadFactoringAsync(Actor, id, null, null, null, 1, false, ct);
        if (!data.TryGetProperty("Lote", out var lot) || lot.ValueKind != JsonValueKind.Object) return NotFound();
        ViewData["IsPrint"] = print;
        return View(lot);
    }

    private FactoringSearch Filter(DateOnly? from, DateOnly? to, string? search, int page = 1)
    {
        var today = Today;
        var term = search?.Trim();
        return new(from ?? today.AddMonths(-3), to ?? today, string.IsNullOrEmpty(term) ? null : term, page);
    }

    private static bool Valid(FactoringSearch filter)
        => filter.From <= filter.To && filter.Page is >= 1 and <= 100000 && filter.Search?.Length is not > 150;

    private Task<JsonElement> ReadAsync(FactoringSearch filter, bool eligible, CancellationToken ct)
        => repository.ReadFactoringAsync(Actor, null, filter.From, filter.To, filter.Search, filter.Page, eligible, ct);
}
