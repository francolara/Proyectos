using System.Security.Claims;
using System.Text.Json;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Billing;
using GestionServicios.Web.ViewModels.Reporting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 10/11/2026 | Listado, registro e impresión de negociaciones de factoring.
// Firma: FRANCO LARA - 12/11/2026 | Imprime el listado de negociaciones sin perder el filtro de fechas ni la búsqueda.
// Firma: FRANCO LARA - 14/11/2026 | Cambios de estado auditados desde la ficha del lote. Filtra los lotes por estado en el listado y su reporte.
[Authorize(Policy = "perm:FACTORING.CONSULTAR")]
[Route("Administracion/Factoring")]
public sealed class FactoringController(IFactoringRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private DateOnly Today => DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
        TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "FACTORING.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(DateOnly? desde, DateOnly? hasta, string? buscar, int pagina = 1, CancellationToken ct = default, bool print = false, string? estado = null)
    {
        var filter = Filter(desde, hasta, buscar, pagina) with { Status = estado };
        if (!ModelState.IsValid || !Valid(filter)) return BadRequest("Revisa los filtros.");
        if (print)
        {
            if (!User.HasClaim(AccessService.PermissionClaim, "REPORTES.IMPRIMIR")) return Forbid();
            try
            {
                var rows = await ListingReport.AllJsonAsync(page => ReadAsync(filter with { Page = page }, false, ct), "LoteId", "Lotes");
                return View("ListingReport", CommercialListingReport.Factoring(rows,
                    [ListingReport.Range(filter.From, filter.To), ListingReport.Filter("Financiera o referencia", filter.Search),
                        ListingReport.Filter("Estado", filter.Status is null ? null : CommercialPage.FactoringStatus(filter.Status))]));
            }
            catch (MaintenanceException error) { return Conflict(error.Message); }
        }
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
        ViewData["CanEdit"] = CanEdit;
        return View(lot);
    }

    [HttpPost("{id:long}/Estado")]
    [Authorize(Policy = "perm:FACTORING.EDITAR")]
    public async Task<IActionResult> ChangeStatus(long id, string estado, CancellationToken ct)
    {
        if (!ModelState.IsValid || id <= 0 || estado is not ("ACEPTADO" or "CERRADO")) return BadRequest();
        try
        {
            await repository.ChangeFactoringStatusAsync(Actor, id, estado, ct);
            TempData["FactoringNotice"] = "Estado del lote actualizado.";
        }
        catch (MaintenanceException error)
        {
            TempData["FactoringError"] = error.Message;
        }
        return RedirectToAction(nameof(Details), new { id });
    }

    private FactoringSearch Filter(DateOnly? from, DateOnly? to, string? search, int page = 1)
    {
        var today = Today;
        var term = search?.Trim();
        return new(from ?? today.AddMonths(-3), to ?? today, string.IsNullOrEmpty(term) ? null : term, page);
    }

    private static bool Valid(FactoringSearch filter)
        => filter.From <= filter.To && filter.Page is >= 1 and <= 100000 && filter.Search?.Length is not > 150 &&
            filter.Status is null or "REGISTRADO" or "ACEPTADO" or "CERRADO";

    private Task<JsonElement> ReadAsync(FactoringSearch filter, bool eligible, CancellationToken ct)
        => repository.ReadFactoringAsync(Actor, null, filter.From, filter.To, filter.Search, filter.Page, eligible, ct, filter.Status);
}
