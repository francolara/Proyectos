using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:PRESUPUESTOS.CONSULTAR")]
[Authorize(Policy = "perm:COSTOS_PROVEEDOR.CONSULTAR")]
[Route("Presupuestos/{id:long}/CostosProveedores")]
public sealed class BudgetSupplierCostsController(ISupplierPaymentsRepository repository, IBudgetRepository budgets) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanRegister => User.HasClaim(AccessService.PermissionClaim, "COSTOS_PROVEEDOR.EDITAR")
        && User.HasClaim(AccessService.PermissionClaim, "PROVEEDORES.CONSULTAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(long id, int page = 1, CancellationToken ct = default)
    {
        if (id <= 0 || page < 1 || !ModelState.IsValid) return BadRequest();
        if (await budgets.GetAsync(Actor, id, null, 1, ct) is null) return NotFound();

        var costs = await repository.CostsAsync(Actor, id, null, null, false, page, ct);
        return View(new BudgetSupplierCostsPage(id, costs, CanRegister, page));
    }

    [HttpGet("Nuevo")]
    [Authorize(Policy = "perm:COSTOS_PROVEEDOR.EDITAR")]
    [Authorize(Policy = "perm:PROVEEDORES.CONSULTAR")]
    public Task<IActionResult> New(long id, bool popup, string? returnUrl,
        CancellationToken ct = default)
        => ShowAsync(id, new SupplierCostForm { OperationId = Guid.NewGuid() }, popup, returnUrl, ct);

    [HttpGet("ElegirProveedor")]
    [Authorize(Policy = "perm:COSTOS_PROVEEDOR.EDITAR")]
    [Authorize(Policy = "perm:PROVEEDORES.CONSULTAR")]
    public async Task<IActionResult> ChooseSupplier(long id, string? search, int page = 1, CancellationToken ct = default)
    {
        if (id <= 0 || page is < 1 or > 100000000 || search?.Length > 180 || !ModelState.IsValid) return BadRequest();
        var suppliers = await repository.SuppliersAsync(Actor, null, search?.Trim(), true, page, ct);
        return PartialView("_ChooseSupplier", new SupplierCostSupplierLookup(id, suppliers, search, page));
    }

    [HttpGet("/Planillas/PagosProveedores/ElegirPresupuesto")]
    [Authorize(Policy = "perm:PAGOS_PROVEEDOR.CONSULTAR")]
    [Authorize(Policy = "perm:COSTOS_PROVEEDOR.EDITAR")]
    [Authorize(Policy = "perm:PROVEEDORES.CONSULTAR")]
    public async Task<IActionResult> ChooseBudget(bool popup, string? search, string? returnUrl,
        int page = 1, CancellationToken ct = default)
    {
        if (page < 1 || page > 100000000 || search?.Length > 180 || !ModelState.IsValid) return BadRequest();

        var results = await budgets.ListAsync(Actor, new DateTimeOffset(1900, 1, 1, 0, 0, 0, TimeSpan.Zero),
            DateTimeOffset.UtcNow.AddDays(1), search?.Trim(), page, null, ct);
        var model = new SupplierCostBudgetLookup(results, search, Back(null, returnUrl));
        return popup ? PartialView("_ChooseBudget", model) : View("ChooseBudget", model);
    }

    [HttpPost("")]
    [Authorize(Policy = "perm:COSTOS_PROVEEDOR.EDITAR")]
    [Authorize(Policy = "perm:PROVEEDORES.CONSULTAR")]
    public async Task<IActionResult> Save(long id, SupplierCostForm input,
        bool popup = false, string? returnUrl = null, CancellationToken ct = default)
    {
        if (id <= 0) return BadRequest();

        if (input.OperationId == Guid.Empty)
            ModelState.AddModelError("", "Vuelve a abrir el presupuesto antes de registrar el costo.");
        if (input.Amount != decimal.Round(input.Amount, 2))
            ModelState.AddModelError(nameof(input.Amount), "Usa hasta dos decimales.");

        if (ModelState.IsValid)
        {
            try
            {
                await repository.SaveCostAsync(Actor,
                    new(input.OperationId, id, input.SupplierId, input.Description, input.Amount,
                        input.Currency, input.Document), ct);
                TempData["SupplierCostSaved"] = "Costo registrado.";
                var destination = Back(id, returnUrl);
                return popup ? Json(new { redirect = destination }) : LocalRedirect(destination);
            }
            catch (SqlException error) when (error.Number is 50000 or 547 or 2601 or 2627)
            {
                ModelState.AddModelError("", "No se pudo registrar el costo. Revisa el proveedor y los datos del concepto.");
            }
        }

        Response.StatusCode = StatusCodes.Status422UnprocessableEntity;
        return await ShowAsync(id, input, popup, returnUrl, ct);
    }

    private string Back(long? id, string? returnUrl)
        => Url.IsLocalUrl(returnUrl) ? returnUrl! : id.HasValue
            ? Url.Action(nameof(Index), new { id })!
            : Url.Action("Index", "SupplierPayments")!;

    private async Task<IActionResult> ShowAsync(long id, SupplierCostForm input,
        bool popup, string? returnUrl, CancellationToken ct)
    {
        if (id <= 0) return BadRequest();

        var budget = await budgets.GetAsync(Actor, id, null, 1, ct);
        if (budget is null) return NotFound();

        var supplier = input.SupplierId > 0
            ? (await repository.SuppliersAsync(Actor, input.SupplierId, null, true, 1, ct)).FirstOrDefault()
            : null;
        if (supplier is null) input.SupplierId = 0;
        var model = new SupplierCostEditor(id, budget.Text("NumeroInterno"), supplier, input, Back(id, returnUrl), popup,
            User.HasClaim(AccessService.PermissionClaim, "OC_CLIENTE.CONSULTAR") ? budget.Text("NumeroOCCliente") : null,
            User.HasClaim(AccessService.PermissionClaim, "GUIAS.CONSULTAR") ? budget.Text("NumeroActa") : null);
        return popup ? PartialView("_Editor", model) : View("New", model);
    }
}

public sealed record BudgetSupplierCostsPage(long BudgetId, IReadOnlyList<SupplierCost> Costs,
    bool CanRegister, int Page);
public sealed record SupplierCostEditor(long BudgetId, string BudgetNumber, Supplier? Supplier,
    SupplierCostForm Input, string ReturnUrl, bool Popup, string? CustomerOrderNumber, string? GuideNumber);
public sealed record SupplierCostSupplierLookup(long BudgetId, IReadOnlyList<Supplier> Suppliers, string? Search, int Page);
public sealed record SupplierCostBudgetLookup(BudgetPage Results, string? Search, string ReturnUrl);

public sealed class SupplierCostForm
{
    public Guid OperationId { get; set; }

    [Range(typeof(long), "1", "9223372036854775807", ErrorMessage = "Selecciona un proveedor.")]
    public long SupplierId { get; set; }

    [Required(ErrorMessage = "Indica la descripción del costo."), StringLength(300)]
    public string Description { get; set; } = "";

    [Range(typeof(decimal), "0.01", "9999999999999999.99", ErrorMessage = "Indica un importe positivo.")]
    public decimal Amount { get; set; }

    [Required, RegularExpression("PEN|USD", ErrorMessage = "Selecciona soles o dólares.")]
    public string Currency { get; set; } = "PEN";

    [StringLength(100)]
    public string? Document { get; set; }
}
