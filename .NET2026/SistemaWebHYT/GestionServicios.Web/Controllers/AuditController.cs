using System.Security.Claims;
using GestionServicios.Application.Security;
using GestionServicios.Web.ViewModels.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:AUDITORIA.CONSULTAR"), Route("Administracion/Auditoria")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class AuditController(ISecurityAdministrationRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private static readonly TimeSpan PeruOffset = TimeSpan.FromHours(-5);

    // Firma: FRANCO LARA - 07/11/2026 | Vincula el filtro de eventos a la consulta y conserva enlaces anteriores.
    [HttpGet]
    public async Task<IActionResult> Index(DateOnly? from, DateOnly? to, string? search, string? actor,
        string? entity, [FromQuery] string? eventAction, [FromQuery(Name = "action")] string? legacyAction,
        int page = 1, CancellationToken ct = default)
    {
        var module = Filter(entity);
        if (module?.ToUpperInvariant() is "PRESUPUESTO" or "PRESUPUESTOREVISION") module = "Presupuesto";
        var action = Filter(Request.Query.ContainsKey(nameof(eventAction)) ? eventAction : legacyAction);
        if (string.Equals(action, nameof(Index), StringComparison.OrdinalIgnoreCase)) action = null;
        var today = DateOnly.FromDateTime(clock.GetUtcNow().ToOffset(PeruOffset).Date);
        var start = from ?? today.AddMonths(-1);
        var end = to ?? today;
        var data = new SecurityAuditPage(new SecurityPage([], 0, 1), []);
        if (end < start || end == DateOnly.MaxValue)
            ModelState.AddModelError("", "La fecha final debe ser igual o posterior a la inicial. Revisa el rango.");
        if (ModelState.IsValid)
            data = await repository.AuditAsync(Actor, Midnight(start), Midnight(end.AddDays(1)),
                Filter(search), Filter(actor), module, action, Math.Clamp(page, 1, 100000000), ct);
        return View(new AuditListModel(data, start, end, search, actor, module, action));
    }

    [HttpGet("{id:long}")]
    public async Task<IActionResult> Detail(long id, CancellationToken ct)
    {
        if (await repository.AuditEventAsync(Actor, id, ct) is not { } row) return NotFound();
        if (Request.Headers["X-Requested-With"] == "XMLHttpRequest") return PartialView("_Detail", row);
        return View(row);
    }

    private static DateTimeOffset Midnight(DateOnly day) => new(day.ToDateTime(TimeOnly.MinValue), PeruOffset);
    private static string? Filter(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
