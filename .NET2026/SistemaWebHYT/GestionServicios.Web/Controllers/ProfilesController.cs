using System.Security.Claims;
using System.Text.Json;
using GestionServicios.Application.Maintenance;
using GestionServicios.Application.Security;
using GestionServicios.Web.ViewModels.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 07/11/2026 | Consulta, edición y duplicación de perfiles por módulo.
[Authorize(Policy = "perm:PERFILES.CONSULTAR"), Route("Administracion/Perfiles")]
public sealed class ProfilesController(ISecurityAdministrationRepository repository) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;

    [HttpGet]
    public async Task<IActionResult> Index(string? search, int page = 1, CancellationToken ct = default) =>
        View(new ProfileListModel(await repository.RolesAsync(Actor, search?.Trim(), Math.Clamp(page, 1, 100000000), ct), search));

    [HttpGet("Nuevo"), Authorize(Policy = "perm:PERFILES.EDITAR")]
    public Task<IActionResult> New(CancellationToken ct) => Form(new(), null, false, ct);

    [HttpGet("{id}")]
    public Task<IActionResult> Detail(string id, CancellationToken ct) => Open(id, true, false, ct);

    [HttpGet("{id}/Editar"), Authorize(Policy = "perm:PERFILES.EDITAR")]
    public Task<IActionResult> Edit(string id, CancellationToken ct) => Open(id, false, false, ct);

    [HttpGet("{id}/Duplicar"), Authorize(Policy = "perm:PERFILES.EDITAR")]
    public Task<IActionResult> Duplicate(string id, CancellationToken ct) => Open(id, false, true, ct);

    [HttpPost("Nuevo"), Authorize(Policy = "perm:PERFILES.EDITAR")]
    public Task<IActionResult> Create(ProfileInput input, CancellationToken ct) => Save(null, input, ct);

    [HttpPost("{id}/Editar"), Authorize(Policy = "perm:PERFILES.EDITAR")]
    public Task<IActionResult> Update(string id, ProfileInput input, CancellationToken ct) => Save(id, input, ct);

    private async Task<IActionResult> Open(string id, bool readOnly, bool duplicate, CancellationToken ct)
    {
        if (await repository.RoleAsync(Actor, id, ct) is not { } row) return NotFound();
        var permissions = row.TryGetProperty("Permisos", out var codes) && codes.ValueKind == JsonValueKind.Array
            ? codes.EnumerateArray().Select(code => code.Text("Codigo")).ToArray() : [];
        return await Form(new ProfileInput
        {
            Name = duplicate ? "Copia de " + row.Text("Name") : row.Text("Name"),
            Version = duplicate ? null : row.Text("ConcurrencyStamp", ""),
            Permissions = permissions
        }, duplicate ? null : row, readOnly, ct);
    }

    private async Task<IActionResult> Save(string? id, ProfileInput input, CancellationToken ct)
    {
        var row = id is null ? null : await repository.RoleAsync(Actor, id, ct);
        if (id is not null && row is null) return NotFound();
        var catalog = await VisiblePermissions(ct);
        var codes = catalog.Select(p => p.Text("Codigo")).ToHashSet(StringComparer.Ordinal);
        if (input.Permissions.Any(p => !codes.Contains(p)))
            ModelState.AddModelError("", "Revisa los permisos seleccionados.");
        if (id is not null && string.IsNullOrWhiteSpace(input.Version))
            ModelState.AddModelError("", "Vuelve a abrir el perfil antes de editarlo.");
        if (ModelState.IsValid)
        {
            try
            {
                var roleId = id ?? Guid.NewGuid().ToString("N");
                await repository.SaveRoleAsync(Actor, new(roleId, id is null ? null : input.Version,
                    input.Name.Trim(), input.Permissions.Distinct(StringComparer.Ordinal).ToArray()), ct);
                TempData["SecurityNotice"] = "Perfil guardado.";
                return RedirectToAction(nameof(Detail), new { id = roleId });
            }
            catch (MaintenanceException error) { ModelState.AddModelError("", error.Message); }
        }
        return View("Form", new ProfileFormModel(input, catalog, row, false));
    }

    private async Task<IActionResult> Form(ProfileInput input, JsonElement? record, bool readOnly, CancellationToken ct) =>
        View("Form", new ProfileFormModel(input, await VisiblePermissions(ct), record, readOnly));

    private async Task<IReadOnlyList<JsonElement>> VisiblePermissions(CancellationToken ct)
    {
        var catalog = (await repository.CatalogsAsync(Actor, ct)).Permissions;
        var fineModules = catalog.Select(p => p.Text("Codigo"))
            .Where(code => code.EndsWith(".MODIFICAR", StringComparison.Ordinal))
            .Select(code => code.Split('.')[0]).ToHashSet(StringComparer.Ordinal);
        return catalog.Where(p =>
        {
            var code = p.Text("Codigo");
            return !code.EndsWith(".EDITAR", StringComparison.Ordinal) || !fineModules.Contains(code.Split('.')[0]);
        }).ToArray();
    }
}
