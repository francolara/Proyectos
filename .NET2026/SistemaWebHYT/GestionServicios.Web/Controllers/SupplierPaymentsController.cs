using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Reporting;
using GestionServicios.Application.Maintenance;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 11/11/2026 | Imprime los trabajos de proveedores con totales por moneda y filtros del listado.
[Authorize(Policy = "perm:PAGOS_PROVEEDOR.CONSULTAR")]
[Route("Planillas/PagosProveedores")]
public sealed class SupplierPaymentsController(ISupplierPaymentsRepository repository) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "PAGOS_PROVEEDOR.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(long? budgetId, long? supplierId, bool all = false,
        int page = 1, DateOnly? from = null, DateOnly? to = null, string? status = null,
        string? search = null, CancellationToken ct = default, bool print = false)
    {
        status ??= all ? "TODOS" : "PENDIENTES";
        if (!ModelState.IsValid || page < 1 || budgetId is <= 0 || supplierId is <= 0
            || from > to || to == DateOnly.MaxValue || search?.Length > 180
            || status is not ("TODOS" or "PAGADOS" or "PENDIENTES"))
            return BadRequest();

        var filters = new SupplierCostFilters(from, to, status, search?.Trim());
        Task<IReadOnlyList<SupplierCost>> Load(int number) => repository.CostsAsync(Actor, budgetId, supplierId, null, false, number, ct, filters);
        if (print)
        {
            if (!User.HasClaim(AccessService.PermissionClaim, "REPORTES.IMPRIMIR")) return Forbid();
            try
            {
                var rows = await ListingReport.AllAsync(async number =>
                {
                    var result = await Load(number);
                    return (result, result.FirstOrDefault()?.Total ?? 0);
                }, (SupplierCost row) => row.Id);
                var summary = new List<string> { ListingReport.Range(from, to), ListingReport.Filter("Buscar", filters.Search),
                    ListingReport.Filter("Estado", status) };
                if (budgetId.HasValue) summary.Add(ListingReport.Filter("Presupuesto", rows.FirstOrDefault()?.BudgetNumber ?? budgetId.ToString()));
                if (supplierId.HasValue) summary.Add(ListingReport.Filter("Proveedor", rows.FirstOrDefault()?.SupplierName ?? supplierId.ToString()));
                return View("ListingReport", new ListingReport("Trabajos de proveedores", summary, rows.Count,
                    new(null, ["Fecha", "Proveedor", "Presupuesto", "Trabajo", "Documento de sustento / factura", "Moneda", "Importe", "Pagado", "Saldo", "Estado"],
                        rows.Select(row => new[] { ListingReport.Date(row.CostDate), row.SupplierName, row.BudgetNumber, row.Description,
                            string.Join("\n", new[] { row.Document, row.PurchaseInvoice }.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()), row.Currency, ListingReport.Money(row.Amount), ListingReport.Money(row.Paid),
                            ListingReport.Money(row.Balance), row.Status })),
                    new("Totales por moneda", ["Moneda", "Importe", "Pagado", "Saldo"],
                        rows.GroupBy(row => row.Currency).Select(group => new[] { group.Key, ListingReport.Money(group.Sum(row => row.Amount)),
                            ListingReport.Money(group.Sum(row => row.Paid)), ListingReport.Money(group.Sum(row => row.Balance)) }))));
            }
            catch (MaintenanceException error) { return Conflict(error.Message); }
        }
        var costs = await Load(page);
        return View(new SupplierPaymentsList(costs, budgetId, supplierId, filters, page));
    }

    [HttpGet("/Planillas/PagosPendientesProveedores")]
    public async Task<IActionResult> Pending(int page = 1, DateOnly? from = null, DateOnly? to = null,
        string? search = null, CancellationToken ct = default)
    {
        if (!ModelState.IsValid || page < 1 || from > to || to == DateOnly.MaxValue || search?.Length > 180)
            return BadRequest();

        var filters = new SupplierCostFilters(from, to, "PENDIENTES", search?.Trim());
        var costs = await repository.CostsAsync(Actor, null, null, null, false, page, ct, filters);
        return View("Index", new SupplierPaymentsList(costs, null, null, filters, page, PendingQueue: true, CanPay: CanEdit));
    }

    [HttpGet("{id:long}/RegistrarPago")]
    [Authorize(Policy = "perm:PAGOS_PROVEEDOR.EDITAR")]
    [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
    public Task<IActionResult> PayPending(long id, bool popup, string? returnUrl, CancellationToken ct)
        => ShowPendingAsync(id, new PaymentForm { OperationId = Guid.NewGuid() }, popup, returnUrl, ct);

    [HttpGet("{id:long}")]
    [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
    public Task<IActionResult> Details(long id, int page = 1, CancellationToken ct = default)
        => ShowAsync(id, page, new PaymentForm { OperationId = Guid.NewGuid() }, ct);

    [HttpPost("{id:long}")]
    [Authorize(Policy = "perm:PAGOS_PROVEEDOR.EDITAR")]
    [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
    public async Task<IActionResult> Save(long id, PaymentForm input, CancellationToken ct,
        bool queue = false, bool popup = false, string? returnUrl = null)
    {
        if (input.OperationId == Guid.Empty)
            ModelState.AddModelError("", "Vuelve a abrir el costo antes de registrar el abono.");
        if (input.Amount != decimal.Round(input.Amount, 2))
            ModelState.AddModelError(nameof(input.Amount), "Usa hasta dos decimales.");
        if (ModelState.IsValid)
        {
            try
            {
                await repository.SavePaymentAsync(Actor,
                    new(input.OperationId, id, input.Amount, input.Date!.Value, input.Reference,
                        input.Frequency, null, null), ct);
                TempData["SupplierPaymentSaved"] = "Abono registrado.";
                if (!queue) return RedirectToAction(nameof(Details), new { id });
                var destination = PendingReturn(returnUrl);
                return popup ? Json(new { redirect = destination }) : LocalRedirect(destination);
            }
            catch (SqlException error) when (error.Number == 50000)
            {
                ModelState.AddModelError("", error.Message == "El importe excede el saldo pendiente."
                    ? "El importe excede el saldo pendiente. Revisa el saldo actualizado."
                    : "No se pudo registrar el abono. Revisa los datos e inténtalo nuevamente.");
            }
        }

        if (!queue) return await ShowAsync(id, 1, input, ct);
        Response.StatusCode = StatusCodes.Status422UnprocessableEntity;
        return await ShowPendingAsync(id, input, popup, returnUrl, ct);
    }

    private string PendingReturn(string? url)
        => Url.IsLocalUrl(url) ? url! : Url.Action(nameof(Pending))!;

    private async Task<IActionResult> ShowPendingAsync(long id, PaymentForm input, bool popup,
        string? returnUrl, CancellationToken ct)
    {
        if (id <= 0) return BadRequest();
        var costs = await repository.CostsAsync(Actor, null, null, id, false, 1, ct);
        if (costs.Count == 0) return NotFound();

        var model = new SupplierPaymentDetails(costs[0], [], input, CanEdit, 1, true, popup, PendingReturn(returnUrl));
        return popup ? PartialView("_PayPending", model) : View("PayPending", model);
    }

    private async Task<IActionResult> ShowAsync(long id, int page, PaymentForm input, CancellationToken ct)
    {
        if (id <= 0 || page < 1) return BadRequest();

        var costs = await repository.CostsAsync(Actor, null, null, id, false, 1, ct);
        if (costs.Count == 0) return NotFound();

        var payments = await repository.PaymentsAsync(Actor, id, page, ct);
        return View("Details", new SupplierPaymentDetails(costs[0], payments, input, CanEdit, page));
    }
}

public sealed record SupplierPaymentsList(IReadOnlyList<SupplierCost> Costs, long? BudgetId,
    long? SupplierId, SupplierCostFilters Filters, int Page, bool PendingQueue = false, bool CanPay = false);

public sealed record SupplierPaymentDetails(SupplierCost Cost, IReadOnlyList<SupplierPayment> Payments,
    PaymentForm Input, bool CanEdit, int Page, bool Queue = false, bool Popup = false, string? ReturnUrl = null);

public sealed class PaymentForm
{
    public Guid OperationId { get; set; }

    [Range(typeof(decimal), "0.01", "9999999999999999.99", ErrorMessage = "Indica un importe positivo.")]
    public decimal Amount { get; set; }

    [Required(ErrorMessage = "Indica la fecha del abono.")]
    public DateOnly? Date { get; set; }

    [Required(ErrorMessage = "Indica la referencia del abono."), StringLength(100)]
    public string Reference { get; set; } = "";

    [RegularExpression("SEMANAL|QUINCENAL|MENSUAL", ErrorMessage = "Selecciona una frecuencia válida.")]
    public string? Frequency { get; set; }
}
