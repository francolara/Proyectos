using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:PAGOS_PERSONAL.CONSULTAR")]
[Route("Planillas/Trabajadores/{id:long}/Pagos")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class PersonalPaymentsController(IPersonalPaymentsRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private DateOnly Today => DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
        TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);

    [HttpGet("")]
    public async Task<IActionResult> Index(long id, string? month, CancellationToken ct)
    {
        var today = Today;
        var period = new DateOnly(today.Year, today.Month, 1);
        if (id <= 0 || month is not null && !TryMonth(month, out period)) return BadRequest();

        var account = await repository.AccountAsync(Actor, id, period, ct);
        if (account is null) return NotFound();

        var date = today >= period && today <= period.AddDays(DateTime.DaysInMonth(period.Year, period.Month) - 1)
            ? today : period;
        return View(new PersonalPaymentsPage(account, period.ToString("yyyy-MM", CultureInfo.InvariantCulture),
            date, Guid.NewGuid(), Has("EDITAR")));
    }

    [HttpPost("Movimiento")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.EDITAR")]
    public async Task<IActionResult> Save(long id, string month, PersonalMovementForm input, CancellationToken ct)
    {
        if (id <= 0 || !TryMonth(month, out var period)) return BadRequest();
        if (!ModelState.IsValid || input.OperationId == Guid.Empty || input.Amount != decimal.Round(input.Amount, 2)
            || input.Date!.Value.Year != period.Year || input.Date.Value.Month != period.Month
            || (input.Frequency is null ? input.PeriodFrom.HasValue || input.PeriodTo.HasValue
                : !input.PeriodFrom.HasValue || !input.PeriodTo.HasValue || input.PeriodFrom > input.PeriodTo))
            return Finish(id, month, "Revisa el concepto, el importe y las fechas del movimiento.", true);

        try
        {
            await repository.RegisterMovementAsync(Actor, new(input.OperationId, id, input.Concept,
                input.Date.Value, input.Description, input.Amount, input.Currency, input.Frequency,
                input.PeriodFrom, input.PeriodTo, input.Reference, input.TerminationId), ct);
            return Finish(id, month, "Movimiento registrado.");
        }
        catch (SqlException error) when (error.Number is 50000 or 547 or 2601 or 2627)
        {
            var message = error.Message switch
            {
                "El importe excede el saldo del trabajador a esa fecha." => "El importe excede el saldo a la fecha indicada.",
                "Indique la referencia del pago o adelanto." => "Indica la referencia del pago o adelanto.",
                "Seleccione un cese valido del trabajador." => "Selecciona el cese al que corresponde el concepto.",
                "El movimiento corresponde a un mes cerrado o anterior al ultimo cierre." => "El movimiento corresponde a un mes cerrado.",
                _ => "No se pudo guardar. Revisa el concepto, la fecha y el saldo."
            };
            return Finish(id, month, message, true);
        }
    }

    [HttpGet("Cese")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.CESES")]
    public async Task<IActionResult> Termination(long id, CancellationToken ct)
    {
        if (id <= 0) return BadRequest();
        var today = Today;
        var account = await repository.AccountAsync(Actor, id, new DateOnly(today.Year, today.Month, 1), ct);
        return account is null ? NotFound() : View(new WorkerTerminationPage(account, today, Guid.NewGuid()));
    }

    [HttpPost("Cese")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.CESES")]
    public async Task<IActionResult> Terminate(long id, PersonalTerminationForm input, CancellationToken ct)
    {
        if (id <= 0) return BadRequest();
        if (!ModelState.IsValid || input.OperationId == Guid.Empty || input.LastWorked > input.Date || input.Date > Today)
            return FinishTermination(id, "Revisa la fecha del cese y el último día trabajado.", true);
        try
        {
            await repository.RegisterTerminationAsync(Actor,
                new(input.OperationId, id, input.Date!.Value, input.LastWorked!.Value, input.Reason), ct);
            return FinishTermination(id, "Cese registrado.");
        }
        catch (SqlException error) when (error.Number is 50000 or 2601 or 2627)
        {
            return FinishTermination(id, "No se pudo registrar el cese. Revisa el estado del trabajador y las fechas.", true);
        }
    }

    [HttpPost("Cese/Reingreso")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.CESES")]
    public async Task<IActionResult> Reenter(long id, WorkerReentryForm input, CancellationToken ct)
    {
        if (id <= 0) return BadRequest();
        if (!ModelState.IsValid || input.OperationId == Guid.Empty || input.TerminationId <= 0 || input.Date > Today)
            return FinishTermination(id, "Revisa la fecha de reingreso.", true);

        try
        {
            await repository.RegisterReentryAsync(Actor, input.OperationId, id, input.TerminationId,
                input.Date!.Value, input.Reason, ct);
            return FinishTermination(id, "Reingreso registrado. El trabajador está activo.");
        }
        catch (SqlException error) when (error.Number is 50000 or 2601 or 2627)
        {
            return FinishTermination(id, "Verifica que el cese siga vigente y que el reingreso sea posterior al cese.", true);
        }
    }

    private RedirectToActionResult FinishTermination(long id, string message, bool error = false)
    {
        TempData[error ? "TerminationError" : "TerminationMessage"] = message;
        return RedirectToAction(nameof(Termination), new { id });
    }

    private bool Has(string action) => User.HasClaim(AccessService.PermissionClaim, "PAGOS_PERSONAL." + action);

    private static bool TryMonth(string? month, out DateOnly result)
        => DateOnly.TryParseExact(month + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);

    private RedirectToActionResult Finish(long id, string month, string message, bool error = false)
    {
        TempData[error ? "PersonalPaymentError" : "PersonalPaymentMessage"] = message;
        return RedirectToAction(nameof(Index), new { id, month });
    }
}

public sealed record PersonalPaymentsPage(PersonalAccount Account, string Month, DateOnly Date,
    Guid MovementOperation, bool CanEdit);

public sealed record WorkerTerminationPage(PersonalAccount Account, DateOnly Today, Guid OperationId);

public sealed class PersonalMovementForm
{
    public Guid OperationId { get; set; }
    [Required, RegularExpression("REMUNERACION|SERVICIO|GRATIFICACION|LIQUIDACION|VACACIONES_PENDIENTES|ADELANTO|ADELANTO_GRATIFICACION|DESCUENTO|PAGO")]
    public string Concept { get; set; } = "";
    [Required] public DateOnly? Date { get; set; }
    [Required, StringLength(300)] public string Description { get; set; } = "";
    [Range(typeof(decimal), "0.01", "9999999999999999.99")] public decimal Amount { get; set; }
    [Required, RegularExpression("PEN|USD")] public string Currency { get; set; } = "";
    [RegularExpression("SEMANAL|QUINCENAL|MENSUAL")] public string? Frequency { get; set; }
    public DateOnly? PeriodFrom { get; set; }
    public DateOnly? PeriodTo { get; set; }
    [StringLength(100)] public string? Reference { get; set; }
    public long? TerminationId { get; set; }
}

public sealed class PersonalTerminationForm
{
    public Guid OperationId { get; set; }
    [Required] public DateOnly? Date { get; set; }
    [Required] public DateOnly? LastWorked { get; set; }
    [StringLength(300)] public string? Reason { get; set; }
}

public sealed class WorkerReentryForm
{
    public Guid OperationId { get; set; }
    public long TerminationId { get; set; }
    [Required] public DateOnly? Date { get; set; }
    [StringLength(300)] public string? Reason { get; set; }
}
