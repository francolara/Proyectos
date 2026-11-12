using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Reporting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 11/11/2026 | Reportes de obligaciones mensuales y cobertura SCTR.
[Authorize(Policy = "perm:OBLIGACIONES.CONSULTAR")]
[Route("Planillas/Obligaciones")]
public sealed class PayrollObligationsController(IPayrollObligationsRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "OBLIGACIONES.EDITAR");

    [HttpGet("")]
    public Task<IActionResult> Index(string? month, long? obligationId, CancellationToken ct, bool print = false)
        => obligationId.HasValue
            ? Task.FromResult<IActionResult>(RedirectToAction(nameof(ObligationDetails), new { month, obligationId }))
            : ShowAsync(month, null, null, print ? "ListingReport" : "Index", ct);

    [HttpGet("Ficha/{obligationId:long?}")]
    [Authorize(Policy = "perm:OBLIGACIONES.EDITAR")]
    public Task<IActionResult> ObligationDetails(string? month, long? obligationId, CancellationToken ct)
        => ShowAsync(month, obligationId, null, "ObligationDetails", ct);

    [HttpGet("Tasas")]
    [Authorize(Policy = "perm:OBLIGACIONES.EDITAR")]
    public async Task<IActionResult> Rates(string month, CancellationToken ct)
    {
        if (!TryMonth(month, out var period)) return BadRequest();
        var rates = await repository.PensionEstimatesAsync(Actor, period, ct);
        return View(new PensionRatesPage(month, rates));
    }

    [HttpPost("Tasas")]
    [Authorize(Policy = "perm:OBLIGACIONES.EDITAR")]
    public async Task<IActionResult> SaveRate(string month, PensionRateForm input, CancellationToken ct)
    {
        if (!TryMonth(month, out _)) return BadRequest();
        if (!ModelState.IsValid || input.Rate != decimal.Round(input.Rate, 4))
            TempData["ObligationError"] = "Ingresa un porcentaje mayor que cero, hasta 100 y con un máximo de cuatro decimales.";
        else
        {
            try
            {
                await repository.SavePensionRateAsync(Actor, input.Entity, input.Rate, ct);
                TempData["ObligationMessage"] = "Porcentaje guardado.";
            }
            catch (SqlException error) when (error.Number == 50000)
            {
                TempData["ObligationError"] = "No se pudo guardar el porcentaje de la entidad.";
            }
        }
        return RedirectToAction(nameof(Rates), new { month });
    }

    [HttpGet("/Planillas/SCTR")]
    public async Task<IActionResult> Sctr(string? estado, string? trabajador, CancellationToken ct, bool print = false)
    {
        if (!ModelState.IsValid || estado is not (null or "VIGENTE" or "POR_VENCER" or "VENCIDA" or "PROGRAMADA") ||
            trabajador?.Length > 150) return BadRequest("Revisa los filtros de SCTR.");

        if (print && !User.HasClaim(AccessService.PermissionClaim, "REPORTES.IMPRIMIR")) return Forbid();
        var today = TimeZoneInfo.ConvertTime(clock.GetUtcNow(), TimeZoneInfo.FindSystemTimeZoneById("America/Lima"));
        var data = await repository.ObligationsAsync(Actor, new DateOnly(today.Year, today.Month, 1), ct);
        var search = trabajador?.Trim() ?? "";
        var comparer = CultureInfo.GetCultureInfo("es-PE").CompareInfo;
        var policies = data.Polizas.Where(policy =>
            (estado is null || policy.Estado == estado) &&
            (search.Length == 0 || policy.Trabajadores?.Any(worker =>
                comparer.IndexOf(worker.Nombre, search, CompareOptions.IgnoreCase | CompareOptions.IgnoreNonSpace) >= 0) == true))
            .ToArray();
        if (print)
        {
            bool Matches(SctrWorker worker) => search.Length > 0 &&
                comparer.IndexOf(worker.Nombre, search, CompareOptions.IgnoreCase | CompareOptions.IgnoreNonSpace) >= 0;
            return View("ListingReport", new ListingReport("SCTR", [ListingReport.Filter("Estado", estado),
                ListingReport.Filter("Trabajador", search)], policies.Length,
                new(null, ["Aseguradora", "Póliza", "Cobertura", "Desde", "Hasta", "Estado", "Trabajadores"],
                    policies.Select(row => new[] { row.Aseguradora, row.NumeroPoliza, row.Cobertura, ListingReport.Date(row.VigenciaDesde),
                        ListingReport.Date(row.VigenciaHasta), row.Estado.Replace('_', ' '), (row.Trabajadores?.Length ?? 0).ToString() })),
                new("Trabajadores cubiertos", ["Aseguradora", "Póliza", "Código de trabajador", "Trabajador", "Coincide con búsqueda"],
                    policies.SelectMany(row => (row.Trabajadores ?? []).Select(worker => new[] { row.Aseguradora, row.NumeroPoliza,
                        worker.TrabajadorId.ToString(), worker.Nombre, search.Length == 0 ? "" : Matches(worker) ? "Sí" : "No" })))));
        }
        return View(new SctrList(policies, CanEdit, estado, search));
    }

    [HttpGet("/Planillas/SCTR/Ficha/{policyId:long?}")]
    [Authorize(Policy = "perm:OBLIGACIONES.EDITAR")]
    public Task<IActionResult> SctrForm(long? policyId, CancellationToken ct)
        => ShowAsync(null, null, policyId, "SctrForm", ct);

    private async Task<IActionResult> ShowAsync(string? month, long? obligationId, long? policyId, string view, CancellationToken ct)
    {
        var today = TimeZoneInfo.ConvertTime(clock.GetUtcNow(), TimeZoneInfo.FindSystemTimeZoneById("America/Lima"));
        var period = new DateOnly(today.Year, today.Month, 1);
        if (!ModelState.IsValid || month is not null && !TryMonth(month, out period)) return BadRequest();
        if (!CanEdit && (obligationId.HasValue || policyId.HasValue)) return Forbid();

        if (view == "ListingReport" && !User.HasClaim(AccessService.PermissionClaim, "REPORTES.IMPRIMIR")) return Forbid();
        var data = await repository.ObligationsAsync(Actor, period, ct);
        if (view == "ListingReport")
        {
            var rows = data.Obligaciones;
            return View(view, new ListingReport("Obligaciones AFP y ONP",
                [ListingReport.Filter("Mes", period.ToString("MMMM yyyy", CultureInfo.GetCultureInfo("es-PE")))], rows.Length,
                new(null, ["Tipo", "Entidad", "Vencimiento", "Importe (PEN)", "Estado", "Fecha de pago", "Referencia de pago"],
                    rows.Select(row => new[] { row.Tipo, row.Entidad, ListingReport.Date(row.Vencimiento), ListingReport.Money(row.Importe),
                        row.Estado.Replace('_', ' '), ListingReport.Date(row.FechaPago), row.ReferenciaPago ?? "" })),
                new("Totales", ["Moneda", "Pendiente", "Pagado", "Total"],
                    new[] { new[] { "PEN", ListingReport.Money(rows.Where(row => row.FechaPago is null).Sum(row => row.Importe)),
                        ListingReport.Money(rows.Where(row => row.FechaPago is not null).Sum(row => row.Importe)),
                        ListingReport.Money(rows.Sum(row => row.Importe)) } })));
        }
        var obligation = data.Obligaciones.FirstOrDefault(item => item.ObligacionId == obligationId);
        var policy = data.Polizas.FirstOrDefault(item => item.PolizaId == policyId);
        if (obligationId.HasValue && obligation is null || policyId.HasValue && policy is null) return NotFound();
        var estimates = view == "ObligationDetails"
            ? await repository.PensionEstimatesAsync(Actor, period, ct, includeWorkers: true) : [];
        return View(view, new PayrollObligationsPage(data, period.ToString("yyyy-MM", CultureInfo.InvariantCulture),
            obligation, policy, CanEdit, estimates));
    }

    [HttpPost("AFP-ONP")]
    [Authorize(Policy = "perm:OBLIGACIONES.EDITAR")]
    public async Task<IActionResult> SaveObligation(string month, ObligationForm input, CancellationToken ct)
    {
        if (!TryMonth(month, out var period) || !TryVersion(input.Id, input.Version, out var version)) return BadRequest();
        if (!ModelState.IsValid || input.Amount != decimal.Round(input.Amount, 2))
            return Finish(month, "Revisa la entidad, el importe y el vencimiento.", true, obligationId: input.Id);

        try
        {
            await repository.SaveObligationAsync(Actor, new(input.Id, input.Type, input.Entity, period,
                input.Amount, input.Due!.Value, null, null, version, ""), ct, preservePayment: true);
            return Finish(month, "Obligación guardada.");
        }
        catch (SqlException error) when (error.Number is 50000 or 2601 or 2627)
        {
            return Finish(month, error.Number == 50000
                ? "No se pudo guardar la obligación. Revisa los datos y vuelve a abrir el registro."
                : "Ya existe una obligación para esa entidad, tipo y mes.", true, obligationId: input.Id);
        }
    }

    [HttpPost("{id:long}/ConfirmarPago")]
    [Authorize(Policy = "perm:OBLIGACIONES.EDITAR")]
    public async Task<IActionResult> ConfirmPayment(long id, string month, PaymentConfirmationForm input, CancellationToken ct)
    {
        if (id <= 0 || !TryMonth(month, out var period)) return BadRequest();
        var today = DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
            TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);
        if (!ModelState.IsValid || !TryVersion(id, input.Version, out var version) || input.Date > today)
            return PaymentResult(month, "Revisa la fecha y el número de operación.", true);

        var data = await repository.ObligationsAsync(Actor, period, ct);
        var obligation = data.Obligaciones.FirstOrDefault(row => row.ObligacionId == id);
        if (obligation is null) return NotFound();
        if (obligation.FechaPago.HasValue)
            return PaymentResult(month, "La obligación ya está pagada.", true);
        if (!obligation.VersionFila.AsSpan().SequenceEqual(version))
            return PaymentResult(month, "La obligación cambió. Revisa el importe antes de confirmar.", true);

        try
        {
            await repository.SaveObligationAsync(Actor, obligation with
            {
                FechaPago = input.Date,
                ReferenciaPago = input.Reference.Trim()
            }, ct);
            return PaymentResult(month, "Pago de la obligación confirmado.");
        }
        catch (SqlException error) when (error.Number == 50000)
        {
            return PaymentResult(month, "No se pudo confirmar el pago. Revisa si la obligación cambió.", true);
        }
    }

    private RedirectToActionResult PaymentResult(string month, string message, bool error = false)
    {
        TempData[error ? "ObligationError" : "ObligationMessage"] = message;
        return RedirectToAction(nameof(Index), new { month });
    }

    [HttpPost("/Planillas/SCTR")]
    [Authorize(Policy = "perm:OBLIGACIONES.EDITAR")]
    public async Task<IActionResult> SavePolicy(PolicyForm input, CancellationToken ct)
    {
        if (!TryVersion(input.Id, input.Version, out var version)) return BadRequest();
        if (!ModelState.IsValid || input.To < input.From || input.WorkerIds.Length == 0 || input.WorkerIds.Any(id => id <= 0))
            return FinishSctr("Revisa la cobertura, la vigencia y selecciona al menos un trabajador.", true, input.Id);

        try
        {
            var workers = input.WorkerIds.Distinct().Select(id => new SctrWorker(id, "")).ToArray();
            await repository.SavePolicyAsync(Actor,
                new(input.Id, input.Insurer, input.Number, input.Coverage, input.From!.Value, input.To!.Value,
                    version, "", workers), ct);
            return FinishSctr("Póliza guardada.");
        }
        catch (SqlException error) when (error.Number is 50000 or 2601 or 2627)
        {
            return FinishSctr("No se pudo guardar la póliza. Revisa los trabajadores y la vigencia.", true, input.Id);
        }
    }

    private RedirectToActionResult FinishSctr(string message, bool error = false, long policyId = 0)
    {
        TempData[error ? "ObligationError" : "ObligationMessage"] = message;
        return error
            ? RedirectToAction(nameof(SctrForm), new { policyId = policyId == 0 ? (long?)null : policyId })
            : RedirectToAction(nameof(Sctr));
    }

    private RedirectToActionResult Finish(string month, string message, bool error = false, long obligationId = 0)
    {
        TempData[error ? "ObligationError" : "ObligationMessage"] = message;
        return error
            ? RedirectToAction(nameof(ObligationDetails), new { month, obligationId = obligationId == 0 ? (long?)null : obligationId })
            : RedirectToAction(nameof(Index), new { month });
    }

    private static bool TryVersion(long id, string? text, out byte[] version)
    {
        version = [];
        if (id <= 0) return id == 0;
        try { version = Convert.FromBase64String(text ?? ""); }
        catch (FormatException) { return false; }
        return version.Length == 8;
    }

    private static bool TryMonth(string? month, out DateOnly result)
        => DateOnly.TryParseExact(month + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
}

public sealed record SctrList(SctrPolicy[] Policies, bool CanEdit, string? Status, string Worker);

public sealed record PayrollObligationsPage(PayrollObligations Data, string Month,
    PayrollObligation? Obligation, SctrPolicy? Policy, bool CanEdit, IReadOnlyList<PensionEstimate> Estimates);

public sealed record PensionRatesPage(string Month, IReadOnlyList<PensionEstimate> Rates);

public sealed class PensionRateForm
{
    [Required, StringLength(100)] public string Entity { get; set; } = "";
    [Range(typeof(decimal), "0.0001", "100")] public decimal Rate { get; set; }
}

public sealed class ObligationForm
{
    public long Id { get; set; }
    public string? Version { get; set; }
    [Required, RegularExpression("AFP|ONP")] public string Type { get; set; } = "";
    [Required, StringLength(120)] public string Entity { get; set; } = "";
    [Range(typeof(decimal), "0.01", "9999999999999999.99")] public decimal Amount { get; set; }
    [Required] public DateOnly? Due { get; set; }
}

public sealed class PolicyForm
{
    public long Id { get; set; }
    public string? Version { get; set; }
    [Required, StringLength(120)] public string Insurer { get; set; } = "";
    [Required, StringLength(80)] public string Number { get; set; } = "";
    [Required, RegularExpression("SALUD|PENSION")] public string Coverage { get; set; } = "";
    [Required] public DateOnly? From { get; set; }
    [Required] public DateOnly? To { get; set; }
    public long[] WorkerIds { get; set; } = [];
}


public sealed class PaymentConfirmationForm
{
    [Required] public string Version { get; set; } = "";
    [Required] public DateOnly? Date { get; set; }
    [Required, StringLength(100)] public string Reference { get; set; } = "";
}
