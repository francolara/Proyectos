using System.Security.Claims;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:PAGOS_PERSONAL.CONSULTAR")]
[Route("Planillas/PlanificacionPagos")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class PayPlanningController(IPayPlanningRepository repository, IBudgetRepository budgets, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "PAGOS_PERSONAL.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(string frequency = "SEMANAL", DateOnly? from = null, CancellationToken ct = default)
    {
        var now = TimeZoneInfo.ConvertTime(clock.GetUtcNow(), TimeZoneInfo.FindSystemTimeZoneById("America/Lima"));
        var today = DateOnly.FromDateTime(now.DateTime);
        if (!ModelState.IsValid || !ValidFrequency(frequency) || from > new DateOnly(9999, 12, 24)) return BadRequest();

        var period = Start(frequency, from ?? today);
        var plan = await repository.PlanAsync(Actor, frequency, period, ct);
        return View(new PayPlanningPage(plan, CanEdit && plan.FechaPago is null));
    }

    [HttpGet("BuscarPresupuestos")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.EDITAR")]
    [Authorize(Policy = "perm:PRESUPUESTOS.CONSULTAR")]
    public async Task<IActionResult> SearchBudgets(string? search, int page = 1, CancellationToken ct = default)
    {
        if (page < 1 || page > 100000000 || search?.Length > 180 || !ModelState.IsValid) return BadRequest();

        var results = await budgets.ListAsync(Actor, new DateTimeOffset(1900, 1, 1, 0, 0, 0, TimeSpan.Zero),
            clock.GetUtcNow().AddDays(1), search?.Trim(), page, null, ct);
        return PartialView("_BudgetResults", results);
    }

    [HttpPost("")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.EDITAR")]
    public async Task<IActionResult> Save(PayPlanForm input, CancellationToken ct)
    {
        if (!ValidFrequency(input.Frequency) || input.From is not DateOnly from
            || from > new DateOnly(9999, 12, 24) || Start(input.Frequency, from) != from)
            return BadRequest();

        byte[]? version = null;
        if (!string.IsNullOrEmpty(input.Version))
        {
            try { version = Convert.FromBase64String(input.Version); }
            catch (FormatException) { return BadRequest(); }
            if (version.Length != 8) return BadRequest();
        }

        foreach (var line in input.Lines)
        {
            if (!line.AjustesDetallados && line.Ajustes is null) continue;
            line.Ajustes ??= [];
            if (line.Ajustes.Length > 250 || !line.Ajustes.All(ValidAdjustment))
            {
                ModelState.AddModelError("", "Revisa la fecha, el concepto, las horas completas y el importe de cada ajuste.");
                continue;
            }
            line.MinutosExtra = line.Ajustes.Sum(adjustment => adjustment.Horas) * 60;
            line.ImporteExtra = line.Ajustes.Where(adjustment => adjustment.Tipo == "EXTRA").Sum(adjustment => adjustment.Importe);
            line.Descuento = line.Ajustes.Where(adjustment => adjustment.Tipo == "DESCUENTO").Sum(adjustment => adjustment.Importe);
        }

        if (ModelState.IsValid && input.Lines.Count > 0 && input.Lines.All(ValidLine)
            && input.Lines.Select(line => line.TrabajadorId).Distinct().Count() == input.Lines.Count)
        {
            try
            {
                await repository.SavePlanAsync(Actor, input.Frequency, from, version, input.Lines, ct);
                TempData["PayPlanSaved"] = "Planificación guardada.";
                return RedirectToAction(nameof(Index), new { frequency = input.Frequency, from = from.ToString("yyyy-MM-dd") });
            }
            catch (SqlException error) when (error.Number is 50000 or 2601 or 2627)
            {
                ModelState.AddModelError("", "No se pudo guardar la planificación. Revisa si el período o el personal cambió.");
            }
        }
        else
        {
            ModelState.AddModelError("", "Revisa los importes. El descuento debe ser menor o igual al importe base más las horas extra.");
        }

        var plan = await repository.PlanAsync(Actor, input.Frequency, from, ct);
        Response.StatusCode = StatusCodes.Status422UnprocessableEntity;
        var submitted = input.Lines.GroupBy(line => line.TrabajadorId).ToDictionary(group => group.Key, group => group.First());
        if (plan.FechaPago is null && plan.VersionFila.AsSpan().SequenceEqual(version))
        {
            plan = plan with { Filas = plan.Filas.Select(row => submitted.TryGetValue(row.TrabajadorId, out var line)
                ? row with { Base = line.Base, MinutosExtra = line.MinutosExtra, ImporteExtra = line.ImporteExtra,
                    Descuento = line.Descuento, Ajustes = line.Ajustes }
                : row).ToArray() };
        }
        return View("Index", new PayPlanningPage(plan, CanEdit && plan.FechaPago is null));
    }

    [HttpPost("ConfirmarPago")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.EDITAR")]
    public async Task<IActionResult> ConfirmPayment(string frequency, DateOnly from, PaymentConfirmationForm input, CancellationToken ct)
    {
        if (!ValidFrequency(frequency) || from > new DateOnly(9999, 12, 24) || Start(frequency, from) != from)
            return BadRequest();

        byte[] version;
        try { version = Convert.FromBase64String(input.Version ?? ""); }
        catch (FormatException) { return BadRequest(); }
        var today = DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
            TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);
        if (!ModelState.IsValid || version.Length != 8 || input.Date < from || input.Date > today)
            TempData["PayPlanPaymentError"] = "Revisa la fecha de pago y el número de operación.";
        else
        {
            try
            {
                await repository.ConfirmPlanAsync(Actor, frequency, from, version, input.Date!.Value, input.Reference.Trim(), ct);
                TempData["PayPlanSaved"] = "Pago confirmado. La planificación quedó cerrada.";
            }
            catch (SqlException error) when (error.Number is 50000 or 266 or 2601 or 2627)
            {
                TempData["PayPlanPaymentError"] = error.Errors.Cast<SqlError>().FirstOrDefault(item => item.Number == 50000)?.Message
                    ?? "No se pudo confirmar el pago. Revisa el período y los movimientos registrados.";
            }
        }
        return RedirectToAction(nameof(Index), new { frequency, from = from.ToString("yyyy-MM-dd") });
    }

    private static bool ValidFrequency(string frequency) => frequency is "SEMANAL" or "QUINCENAL" or "MENSUAL";

    [HttpGet("Reporte")]
    public async Task<IActionResult> Report(string frequency, DateOnly from, CancellationToken ct)
    {
        if (!ModelState.IsValid || !ValidFrequency(frequency) || from > new DateOnly(9999, 12, 24)
            || Start(frequency, from) != from) return BadRequest();
        var plan = await repository.ReportPlanAsync(Actor, frequency, from, ct);
        return plan.PlanId.HasValue ? View(plan) : NotFound();
    }

    private static bool ValidAdjustment(PayPlanAdjustment adjustment)
        => adjustment.Fecha.HasValue && adjustment.Presupuesto?.Length is null or <= 40
            && adjustment.Concepto?.Length is null or <= 300 && ValidAmount(adjustment.Importe)
            && (adjustment.Tipo == "EXTRA" && adjustment.Horas is >= 1 and <= 24
                && (!string.IsNullOrWhiteSpace(adjustment.Presupuesto) || !string.IsNullOrWhiteSpace(adjustment.Concepto))
                || adjustment.Tipo == "DESCUENTO" && adjustment.Horas == 0 && adjustment.Importe > 0
                && !string.IsNullOrWhiteSpace(adjustment.Concepto));

    private static DateOnly Start(string frequency, DateOnly date) => frequency switch
    {
        "SEMANAL" => date.AddDays(-(((int)date.DayOfWeek + 6) % 7)),
        "QUINCENAL" => new(date.Year, date.Month, date.Day <= 15 ? 1 : 16),
        _ => new(date.Year, date.Month, 1)
    };

    private static bool ValidLine(PayPlanLineInput line)
        => line.TrabajadorId > 0 && line.MinutosExtra >= 0 && ValidAmount(line.Base) && ValidAmount(line.ImporteExtra)
            && ValidAmount(line.Descuento) && line.Descuento <= line.Base + line.ImporteExtra;

    private static bool ValidAmount(decimal amount)
        => amount is >= 0 and <= 9999999999999999.99m && decimal.Round(amount, 2) == amount;
}

public sealed record PayPlanningPage(PayPlan Plan, bool CanEdit);
public sealed record PayPlanAdjustmentField(string Prefix, PayPlanAdjustment? Value, DateOnly From, DateOnly To);

public sealed class PayPlanForm
{
    public string Frequency { get; set; } = "";
    public DateOnly? From { get; set; }
    public string? Version { get; set; }
    public List<PayPlanLineInput> Lines { get; set; } = [];
}
