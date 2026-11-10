using System.Globalization;
using System.Security.Claims;
using System.Text;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:REPORTES_PLANILLAS.CONSULTAR")]
[Route("Planillas/Reporte")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class PayrollReportsController(IPersonalPaymentsRepository repository, IPayrollObligationsRepository obligations, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private DateOnly Today => DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
        TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);

    [HttpGet("")]
    public Task<IActionResult> Index(PayrollReportFilter filter, CancellationToken ct)
        => ShowAsync(filter, false, ct);

    [HttpGet("Imprimir")]
    public Task<IActionResult> Print(PayrollReportFilter filter, CancellationToken ct)
        => ShowAsync(filter, true, ct);

    private async Task<IActionResult> ShowAsync(PayrollReportFilter filter, bool print, CancellationToken ct)
    {
        var today = Today;
        var currentMonth = new DateOnly(today.Year, today.Month, 1);
        var period = currentMonth;
        if (!ModelState.IsValid || filter.SupplierId is <= 0 || filter.BudgetId is <= 0
            || filter.Search?.Length > 180 || filter.Month is not null && !TryMonth(filter.Month, out period))
            return BadRequest();

        var allowed = new List<string>();
        if (Has("PAGOS_PERSONAL.CONSULTAR")) allowed.Add("personal");
        if (Has("PAGOS_PROVEEDOR.CONSULTAR")) allowed.Add("proveedores");
        if (Has("OBLIGACIONES.CONSULTAR")) allowed.Add("obligaciones");
        var kind = filter.Kind ?? (filter.SupplierId.HasValue || filter.BudgetId.HasValue ? "proveedores" : allowed.FirstOrDefault());
        if (kind is null)
            return print ? Forbid() : View(new PayrollReportPage(filter, allowed, null, [], false));
        if (!PayrollReportPage.Titles.ContainsKey(kind) || !ValidFilters(kind, filter)) return BadRequest();
        if (!allowed.Contains(kind)) return Forbid();

        filter.Kind = kind;
        filter.Month = period.ToString("yyyy-MM", CultureInfo.InvariantCulture);
        filter.Search = filter.Search?.Trim();
        MonthlyPayrollReport? report = null;
        PayrollObligation[] rows = [];
        if (kind == "obligaciones")
        {
            var data = await obligations.ObligationsAsync(Actor, period, ct);
            rows = data.Obligaciones.Where(row => (filter.Type is null || row.Tipo == filter.Type)
                && Matches(row.Entidad, filter.Search)
                && (filter.Status is null || (filter.Status == "PENDIENTE" ? row.FechaPago is null : row.Estado == filter.Status))).ToArray();
        }
        else
        {
            report = await repository.ReportAsync(Actor, period, kind == "proveedores" ? filter.SupplierId : null,
                kind == "proveedores" ? filter.BudgetId : null, ct);
            if (kind == "personal" ? !report.VerPersonal : !report.VerProveedores) return Forbid();
            report = FilterRows(report, kind, filter);
        }

        var canClose = kind == "personal" && report is { Cerrado: false } && period < currentMonth
            && Has("PAGOS_PERSONAL.CERRAR");
        return View(print ? "Print" : "Index", new PayrollReportPage(filter, allowed, report, rows, canClose, print));
    }

    private bool Has(string permission) => User.HasClaim(AccessService.PermissionClaim, permission);

    private static bool ValidFilters(string kind, PayrollReportFilter filter) => kind switch
    {
        "personal" => (filter.Type is null or "PLANILLA" or "EXTERNO") && filter.Status is null,
        "proveedores" => filter.Type is null && (filter.Status is null or "PENDIENTE" or "PAGADO"),
        _ => (filter.Type is null or "AFP" or "ONP") && (filter.Status is null or "PENDIENTE" or "PAGADA" or "VENCIDA")
    };

    private static bool Matches(string value, string? search)
        => string.IsNullOrWhiteSpace(search) || CultureInfo.GetCultureInfo("es-PE").CompareInfo.IndexOf(value, search,
            CompareOptions.IgnoreCase | CompareOptions.IgnoreNonSpace) >= 0;

    private static MonthlyPayrollReport FilterRows(MonthlyPayrollReport report, string kind, PayrollReportFilter filter)
        => kind == "personal"
            ? report with { Personal = report.Personal.Where(row => (filter.Type is null || row.TipoTrabajador == filter.Type)
                && Matches(row.NombreCompleto, filter.Search)).ToArray() }
            : report with { Proveedores = report.Proveedores.Where(row =>
                (Matches(row.RazonSocial, filter.Search) || Matches(row.NumeroInterno, filter.Search))
                && (filter.Status is null || (filter.Status == "PENDIENTE" ? row.Pendiente > 0 : row.Pendiente == 0))).ToArray() };

    [HttpGet("Exportar/{kind}")]
    public async Task<IActionResult> Export(string kind, PayrollReportFilter filter, CancellationToken ct)
    {
        if (!ModelState.IsValid || !TryMonth(filter.Month, out var period) || filter.SupplierId is <= 0 || filter.BudgetId is <= 0
            || filter.Search?.Length > 180 || kind is not ("personal" or "proveedores") || !ValidFilters(kind, filter)) return BadRequest();

        var report = await repository.ReportAsync(Actor, period, filter.SupplierId, filter.BudgetId, ct);
        if (kind == "personal" ? !report.VerPersonal : !report.VerProveedores) return Forbid();

        filter.Search = filter.Search?.Trim();
        report = FilterRows(report, kind, filter);
        var csv = new StringBuilder();
        if (kind == "personal")
        {
            csv.AppendLine("Trabajador;Tipo;Moneda;Saldo anterior;Devengado;Adelantado;Descontado;Pagado;Saldo final");
            foreach (var row in report.Personal)
                csv.AppendLine(string.Join(';', Cell(row.NombreCompleto), Cell(row.TipoTrabajador), Cell(row.MonedaCodigo),
                    Amount(row.SaldoAnterior), Amount(row.Devengado), Amount(row.Adelantado), Amount(row.Descontado),
                    Amount(row.Pagado), Amount(row.Saldo)));
        }
        else
        {
            csv.AppendLine("Proveedor;Presupuesto;Moneda;Costo;Pagado en el mes;Pagado acumulado;Pendiente");
            foreach (var row in report.Proveedores)
                csv.AppendLine(string.Join(';', Cell(row.RazonSocial), Cell(row.NumeroInterno), Cell(row.MonedaCodigo),
                    Amount(row.Costo), Amount(row.PagadoMes), Amount(row.PagadoAcumulado), Amount(row.Pendiente)));
        }

        return File(Encoding.UTF8.GetBytes("\uFEFF" + csv), "text/csv; charset=utf-8",
            $"planillas-{kind}-{period.ToString("yyyy-MM", CultureInfo.InvariantCulture)}.csv");
    }

    [HttpPost("Cerrar")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.CONSULTAR")]
    [Authorize(Policy = "perm:PAGOS_PERSONAL.CERRAR")]
    public async Task<IActionResult> Close(string month, bool confirmed, CancellationToken ct)
    {
        if (!confirmed || !TryMonth(month, out var period)) return BadRequest();
        try
        {
            await repository.CloseMonthAsync(Actor, period, ct);
            TempData["PayrollReportMessage"] = "Mes de personal cerrado.";
        }
        catch (SqlException error) when (error.Number == 50000)
        {
            TempData["PayrollReportError"] = error.Message == "Cierre primero los meses anteriores que tienen movimientos."
                ? "Cierra primero los meses anteriores que tienen movimientos."
                : "No se pudo cerrar el mes. Comprueba que el período haya terminado.";
        }
        return RedirectToAction(nameof(Index), new { month, kind = "personal" });
    }

    private static bool TryMonth(string? month, out DateOnly result)
        => DateOnly.TryParseExact(month + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result)
            && result.Year < 9999;

    private static string Amount(decimal value) => value.ToString("0.00", CultureInfo.GetCultureInfo("es-PE"));

    private static string Cell(string value)
    {
        var content = value.TrimStart();
        if (content.Length > 0 && content[0] is '=' or '+' or '-' or '@')
            value = "'" + value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}

public sealed class PayrollReportFilter
{
    public string? Kind { get; set; }
    public string? Month { get; set; }
    public string? Search { get; set; }
    public string? Type { get; set; }
    public string? Status { get; set; }
    public long? SupplierId { get; set; }
    public long? BudgetId { get; set; }

    public Dictionary<string, string> RouteValues()
    {
        var values = new Dictionary<string, string>();
        if (Kind is not null) values["kind"] = Kind;
        if (Month is not null) values["month"] = Month;
        if (!string.IsNullOrWhiteSpace(Search)) values["search"] = Search;
        if (Type is not null) values["type"] = Type;
        if (Status is not null) values["status"] = Status;
        if (SupplierId.HasValue) values["supplierId"] = SupplierId.Value.ToString(CultureInfo.InvariantCulture);
        if (BudgetId.HasValue) values["budgetId"] = BudgetId.Value.ToString(CultureInfo.InvariantCulture);
        return values;
    }
}

public sealed record PayrollReportPage(PayrollReportFilter Filter, IReadOnlyList<string> AllowedKinds,
    MonthlyPayrollReport? Report, PayrollObligation[] Obligations, bool CanClose, bool IsPrint = false)
{
    public static readonly IReadOnlyDictionary<string, string> Titles = new Dictionary<string, string>
    {
        ["proveedores"] = "Pagos a Proveedores",
        ["personal"] = "Planillas",
        ["obligaciones"] = "Obligaciones"
    };

    public string Title => Titles[Filter.Kind!];
    public int Count => Filter.Kind switch
    {
        "personal" => Report!.Personal.Length,
        "proveedores" => Report!.Proveedores.Length,
        _ => Obligations.Length
    };

    public string FilterSummary => string.Join(" · ", new[]
    {
        $"Mes {Filter.Month}",
        Filter.Type is null ? null : $"Tipo {Filter.Type}",
        Filter.Status is null ? null : $"Estado {Filter.Status}",
        string.IsNullOrWhiteSpace(Filter.Search) ? null : $"Búsqueda {Filter.Search}",
        Filter.SupplierId.HasValue && Filter.Kind == "proveedores" ? $"Proveedor ID {Filter.SupplierId}" : null,
        Filter.BudgetId.HasValue && Filter.Kind == "proveedores" ? $"Presupuesto ID {Filter.BudgetId}" : null
    }.OfType<string>());
}
