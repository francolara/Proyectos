using System.Security.Claims;
using System.Text.Json;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Billing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 10/11/2026 | Consulta y pago de detracciones desde el listado de Administración.
[Authorize(Policy = "perm:PAGOS_PROVEEDOR.CONSULTAR")]
[Route("Administracion/Detracciones")]
public sealed class WithholdingPaymentsController(ICommercialDocumentRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanPay => User.HasClaim(AccessService.PermissionClaim, "PAGOS_PROVEEDOR.EDITAR");
    private DateOnly Today => DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
        TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);

    [HttpGet("")]
    public async Task<IActionResult> Index(string estado = "PENDIENTE", string? proveedor = null, string? buscar = null,
        int pagina = 1, CancellationToken ct = default)
    {
        if (!ModelState.IsValid || estado is not ("TODOS" or "PENDIENTE" or "PAGADA") ||
            proveedor?.Length > 150 || buscar?.Length > 150 || pagina is < 1 or > 100000) return BadRequest("Revisa los filtros.");
        var supplier = string.IsNullOrWhiteSpace(proveedor) ? null : proveedor.Trim();
        var search = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim();
        var rows = await repository.ReadWithholdingsAsync(Actor, estado, supplier, search, pagina, ct);
        return View(new WithholdingList(rows, estado, supplier, search, pagina, CanPay));
    }

    [HttpGet("{id:long}/Pagar")]
    [Authorize(Policy = "perm:PAGOS_PROVEEDOR.EDITAR")]
    public async Task<IActionResult> Payment(string? retorno, long id, bool popup, CancellationToken ct)
    {
        if (!popup) return RedirectToAction(nameof(Index));
        var data = await repository.ReadAsync(Actor, false, id, null, null, null, null, 1, null, null, ct);
        if (!data.TryGetProperty("Documento", out var record) || record.ValueKind != JsonValueKind.Object) return NotFound();
        var paid = record.GetProperty("Pagos").EnumerateArray().Where(row => row.GetProperty("EsDetraccion").GetBoolean())
            .Sum(row => CommercialPage.Amount(row, "Importe"));
        var balance = CommercialPage.Amount(record, "ImporteDetraccion") - paid;
        if (balance <= 0 || CommercialPage.Text(record, "Tipo") == "07") return BadRequest("La detracción no tiene saldo pendiente.");
        return PartialView("_Payment", new WithholdingPayment(record, balance, Today, ReturnUrl(retorno)));
    }

    [HttpPost("{id:long}/Pagar")]
    [Authorize(Policy = "perm:PAGOS_PROVEEDOR.EDITAR")]
    public async Task<IActionResult> Pay(long id, DateOnly fecha, decimal importe, string? operacion, string? retorno, CancellationToken ct)
    {
        if (!ModelState.IsValid || fecha == default || fecha > Today || importe <= 0 ||
            string.IsNullOrWhiteSpace(operacion) || operacion.Length > 100)
            return BadRequest("Completa la fecha, el importe y el número de operación.");
        try
        {
            await repository.PayAsync(Actor, id, true, fecha, importe, operacion.Trim(), ct);
            TempData["WithholdingNotice"] = "Pago de detracción registrado.";
            return Json(new { redirect = ReturnUrl(retorno) });
        }
        catch (MaintenanceException error) { return BadRequest(error.Message); }
    }

    private string ReturnUrl(string? url) => Url.IsLocalUrl(url) ? url! : Url.Action(nameof(Index))!;
}

public sealed record WithholdingList(JsonElement Rows, string State, string? Supplier, string? Search, int Page, bool CanPay);
public sealed record WithholdingPayment(JsonElement Document, decimal Balance, DateOnly Today, string ReturnUrl);
