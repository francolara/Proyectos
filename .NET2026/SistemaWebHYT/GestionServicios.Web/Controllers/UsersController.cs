using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Json;
using GestionServicios.Application.Maintenance;
using GestionServicios.Application.Security;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 07/11/2026 | Administración de usuarios, contraseñas temporales y sesiones.
[Authorize(Policy = "perm:USUARIOS.CONSULTAR"), Route("Administracion/Usuarios")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class UsersController(ISecurityAdministrationRepository repository, IPasswordHasher<AccessAccount> hasher) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;

    private static readonly IReadOnlyDictionary<string, string> ActionPermissions = new Dictionary<string, string>
    {
        ["ACTIVAR"] = "ACTIVAR", ["DESACTIVAR"] = "ACTIVAR", ["RESTABLECER"] = "RESTABLECER",
        ["DESBLOQUEAR"] = "DESBLOQUEAR", ["CERRAR_SESIONES"] = "SESIONES"
    };

    [HttpGet]
    public async Task<IActionResult> Index(string? search, bool? active, string? situation, int page = 1, CancellationToken ct = default)
    {
        situation = situation is "TEMPORAL" or "BLOQUEADO" or "DISPONIBLE" ? situation : null;
        var result = await repository.UsersAsync(Actor, search?.Trim(), active, situation, Math.Clamp(page, 1, 100000000), ct);
        return View(new UserListModel(result, search, active, situation));
    }

    [HttpGet("Nuevo"), Authorize(Policy = "perm:USUARIOS.CREAR")]
    public Task<IActionResult> New(CancellationToken ct) => Form(new() { TemporaryPassword = GeneratePassword() }, null, false, ct);

    [HttpGet("{id}")]
    public Task<IActionResult> Detail(string id, CancellationToken ct) => Open(id, true, ct);

    [HttpGet("{id}/Editar"), Authorize(Policy = "perm:USUARIOS.MODIFICAR")]
    public Task<IActionResult> Edit(string id, CancellationToken ct) => Open(id, false, ct);

    [HttpPost("Nuevo"), Authorize(Policy = "perm:USUARIOS.CREAR")]
    public Task<IActionResult> Create(UserInput input, CancellationToken ct) => Save(null, input, ct);

    [HttpPost("{id}/Editar"), Authorize(Policy = "perm:USUARIOS.MODIFICAR")]
    public Task<IActionResult> Update(string id, UserInput input, CancellationToken ct) => Save(id, input, ct);

    private async Task<IActionResult> Open(string id, bool readOnly, CancellationToken ct)
    {
        if (await repository.UserAsync(Actor, id, ct) is not { } row) return NotFound();
        return await Form(new UserInput
        {
            Names = row.Text("Nombres", ""), Surnames = row.Text("Apellidos", ""), Email = row.Text("Email", ""),
            Active = row.Flag("Activo"), Version = row.Text("ConcurrencyStamp", ""), RoleId = row.Text("PerfilId", "")
        }, row, readOnly, ct);
    }

    private async Task<IActionResult> Save(string? id, UserInput input, CancellationToken ct)
    {
        var row = id is null ? null : await repository.UserAsync(Actor, id, ct);
        if (id is not null && row is null) return NotFound();
        if (id is null && string.IsNullOrWhiteSpace(input.TemporaryPassword))
            ModelState.AddModelError(nameof(input.TemporaryPassword), "Escribe una contraseña temporal.");
        if (id is not null && string.IsNullOrWhiteSpace(input.Version))
            ModelState.AddModelError("", "Vuelve a abrir el registro antes de editarlo.");

        if (ModelState.IsValid)
        {
            try
            {
                var email = input.Email.Trim();
                var targetId = id ?? Guid.NewGuid().ToString("N");
                var hash = id is null ? Hash(targetId, email, input.TemporaryPassword!) : null;
                await repository.SaveUserAsync(Actor, new(targetId, id is null ? null : input.Version,
                    input.Names.Trim(), input.Surnames.Trim(), email, input.RoleId, input.Active, hash), ct);
                if (id is null) return View("TemporaryPassword", new TemporaryPasswordModel(email, input.TemporaryPassword!));
                TempData["SecurityNotice"] = "Usuario actualizado.";
                return RedirectToAction(nameof(Detail), new { id });
            }
            catch (MaintenanceException error) { ModelState.AddModelError("", error.Message); }
        }
        input.TemporaryPassword = id is null ? GeneratePassword() : null;
        ModelState.Remove(nameof(input.TemporaryPassword));
        return await Form(input, row, false, ct);
    }

    [HttpGet("{id}/Accion/{actionCode}")]
    public async Task<IActionResult> Manage(string id, string actionCode, CancellationToken ct)
    {
        if (!ActionPermissions.TryGetValue(actionCode, out var permission)) return NotFound();
        if (!User.HasClaim(AccessService.PermissionClaim, $"USUARIOS.{permission}")) return Forbid();
        if (await repository.UserAsync(Actor, id, ct) is not { } row) return NotFound();
        return View("Manage", ActionModel(row, actionCode, new UserActionInput
        {
            Version = row.Text("ConcurrencyStamp", ""), TemporaryPassword = actionCode == "RESTABLECER" ? GeneratePassword() : null
        }));
    }

    [HttpPost("{id}/Accion/{actionCode}")]
    public async Task<IActionResult> Manage(string id, string actionCode, UserActionInput input, CancellationToken ct)
    {
        if (!ActionPermissions.TryGetValue(actionCode, out var permission)) return NotFound();
        if (!User.HasClaim(AccessService.PermissionClaim, $"USUARIOS.{permission}")) return Forbid();
        if (await repository.UserAsync(Actor, id, ct) is not { } row) return NotFound();
        var email = row.Text("Email", "");
        if (actionCode == "RESTABLECER" && string.IsNullOrWhiteSpace(input.TemporaryPassword))
            ModelState.AddModelError(nameof(input.TemporaryPassword), "Escribe una contraseña temporal.");
        if (ModelState.IsValid)
        {
            try
            {
                var hash = actionCode == "RESTABLECER" ? Hash(id, email, input.TemporaryPassword!) : null;
                await repository.UserActionAsync(Actor, id, input.Version, actionCode, hash, input.Reason.Trim(), ct);
                if (hash is not null) return View("TemporaryPassword", new TemporaryPasswordModel(email, input.TemporaryPassword!));
                TempData["SecurityNotice"] = "Acción registrada.";
                return RedirectToAction(nameof(Detail), new { id });
            }
            catch (MaintenanceException error) { ModelState.AddModelError("", error.Message); }
        }
        input.TemporaryPassword = actionCode == "RESTABLECER" ? GeneratePassword() : null;
        ModelState.Remove(nameof(input.TemporaryPassword));
        return View("Manage", ActionModel(row, actionCode, input));
    }

    private async Task<IActionResult> Form(UserInput input, JsonElement? row, bool readOnly, CancellationToken ct)
    {
        var catalogs = await repository.CatalogsAsync(Actor, ct);
        return View("Form", new UserFormModel(input, catalogs.Roles, row, readOnly));
    }

    private static UserActionModel ActionModel(JsonElement row, string action, UserActionInput input) =>
        new(row.Text("Id"), SecurityDisplay.Name(row), row.Text("Email"), action, input);

    private string Hash(string id, string email, string password) =>
        hasher.HashPassword(new AccessAccount(id, email, email, null, "", true, null, [], []), password);

    private static string GeneratePassword() => Convert.ToHexString(RandomNumberGenerator.GetBytes(12));
}
