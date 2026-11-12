using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 11/11/2026 | Separa consulta y edición del proveedor y conserva el retorno al listado.
[Authorize(Policy = "perm:PROVEEDORES.CONSULTAR")]
[Route("Proveedores")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class SuppliersController(ISupplierPaymentsRepository repository) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "PROVEEDORES.EDITAR");
    private bool CanSeeAccount => User.HasClaim(AccessService.PermissionClaim, "PROVEEDORES.CUENTA");

    [HttpGet("")]
    public async Task<IActionResult> Index(string? search, int page = 1, CancellationToken ct = default)
    {
        if (!ModelState.IsValid || page < 1 || search?.Length > 180) return BadRequest();

        var suppliers = await repository.SuppliersAsync(Actor, null, search, false, page, ct);
        return View(new SuppliersList(suppliers, search, page, CanEdit));
    }

    [HttpGet("Ficha/{id:long?}")]
    public Task<IActionResult> Details(long? id, string? returnUrl, CancellationToken ct)
    {
        if (id is long supplierId) return Load(supplierId, false, returnUrl, ct);
        return Task.FromResult<IActionResult>(CanEdit ? Show(new SupplierForm { Active = true }, true, returnUrl) : Forbid());
    }

    [HttpGet("Ficha/{id:long}/Editar")]
    [Authorize(Policy = "perm:PROVEEDORES.EDITAR")]
    public Task<IActionResult> Edit(long id, string? returnUrl, CancellationToken ct)
        => Load(id, true, returnUrl, ct);

    private async Task<IActionResult> Load(long id, bool editing, string? returnUrl, CancellationToken ct)
    {
        if (id <= 0) return NotFound();
        var records = await repository.SuppliersAsync(Actor, id, null, false, 1, ct);
        if (records.Count == 0) return NotFound();

        var supplier = records[0];
        return Show(new SupplierForm
        {
            Id = supplier.Id, TaxId = supplier.TaxId, Name = supplier.Name, Active = supplier.Active,
            Bank = supplier.Bank, Account = CanSeeAccount ? supplier.Account : supplier.MaskedAccount,
            Cci = CanSeeAccount ? supplier.Cci : null,
            Version = Convert.ToBase64String(supplier.Version)
        }, editing, returnUrl);
    }

    [HttpPost("Ficha/{id:long?}")]
    [Authorize(Policy = "perm:PROVEEDORES.EDITAR")]
    public async Task<IActionResult> Save(long? id, SupplierForm input, string? returnUrl, CancellationToken ct)
    {
        if (id is <= 0) return NotFound();
        input.Id = id;
        byte[]? version = null;
        if (id.HasValue)
        {
            try { version = Convert.FromBase64String(input.Version ?? ""); }
            catch (FormatException) { return BadRequest(); }
            if (version.Length != 8) return BadRequest();
        }

        if (CanSeeAccount && string.IsNullOrWhiteSpace(input.Bank) != string.IsNullOrWhiteSpace(input.Account))
            ModelState.AddModelError("", "Completa el banco y el número de cuenta juntos.");

        if (CanSeeAccount && !string.IsNullOrWhiteSpace(input.Cci) && string.IsNullOrWhiteSpace(input.Bank))
            ModelState.AddModelError(nameof(input.Cci), "Selecciona el banco correspondiente al CCI.");

        if (ModelState.IsValid)
        {
            try
            {
                var savedId = await repository.SaveSupplierAsync(Actor,
                    new(id, input.TaxId, input.Name, input.Active, version, CanSeeAccount,
                        CanSeeAccount ? input.Bank : null, CanSeeAccount ? input.Account : null,
                        CanSeeAccount ? input.Cci : null, CanSeeAccount), ct);
                TempData["SupplierSaved"] = "Proveedor guardado.";
                return RedirectToAction(nameof(Details), new { id = savedId, returnUrl = LocalReturn(returnUrl) });
            }
            catch (SqlException error) when (error.Number is 2601 or 2627)
            {
                ModelState.AddModelError(nameof(input.TaxId), "Ya existe un proveedor con este RUC.");
            }
            catch (SqlException error) when (error.Number == 50000)
            {
                ModelState.AddModelError("", "No se pudo guardar. Vuelve a abrir la ficha y revisa los datos.");
            }
        }

        if (!CanSeeAccount)
        {
            input.Account = null;
            input.Cci = null;
            input.Bank = null;
        }

        return Show(input, true, returnUrl);
    }

    private string? LocalReturn(string? returnUrl) => Url.IsLocalUrl(returnUrl) ? returnUrl : null;

    private ViewResult Show(SupplierForm input, bool editing, string? returnUrl)
        => View("Details", new SupplierDetails(input, CanEdit, CanSeeAccount, editing, LocalReturn(returnUrl)));
}

public sealed record SuppliersList(IReadOnlyList<Supplier> Suppliers, string? Search, int Page, bool CanEdit);
public sealed record SupplierDetails(SupplierForm Input, bool CanEdit, bool CanSeeAccount, bool Editing, string? ReturnUrl);

public sealed class SupplierForm
{
    public long? Id { get; set; }
    public string? Version { get; set; }

    [Required(ErrorMessage = "Indica el RUC.")]
    [RegularExpression("[0-9]{11}", ErrorMessage = "El RUC debe tener once dígitos.")]
    public string TaxId { get; set; } = "";

    [Required(ErrorMessage = "Indica la razón social."), StringLength(180)]
    public string Name { get; set; } = "";

    [StringLength(80)]
    public string? Bank { get; set; }

    [StringLength(34)]
    public string? Account { get; set; }

    [RegularExpression("[0-9]{20}", ErrorMessage = "El CCI debe tener 20 dígitos.")]
    public string? Cci { get; set; }

    public bool Active { get; set; }
}
