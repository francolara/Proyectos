using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:CONFIGURACION.CONSULTAR")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
[Route("Configuracion/HYT/Telecredito")]
public sealed class TelecreditSettingsController(ITelecreditRepository repository) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "CONFIGURACION.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(CancellationToken ct)
    {
        var account = await repository.GetChargeAccountAsync(Actor, ct);
        return View(new TelecreditSettingsForm
        {
            AccountType = account?.TipoCuenta ?? "C",
            AccountNumber = account?.NumeroCuenta,
            Currency = account?.Moneda ?? "PEN",
            Version = account?.VersionFila is byte[] version ? Convert.ToBase64String(version) : null,
            CanEdit = CanEdit
        });
    }

    [HttpPost("")]
    [Authorize(Policy = "perm:CONFIGURACION.EDITAR")]
    public async Task<IActionResult> Save(TelecreditSettingsForm input, CancellationToken ct)
    {
        byte[]? version = null;
        if (!string.IsNullOrEmpty(input.Version))
        {
            try { version = Convert.FromBase64String(input.Version); }
            catch (FormatException) { return BadRequest(); }
            if (version.Length != 8) return BadRequest();
        }

        input.CanEdit = true;
        var account = Telecredit.NormalizeAccount(input.AccountNumber);
        if (!Telecredit.HasChargeAccountFormat(account))
            ModelState.AddModelError(nameof(input.AccountNumber), "La cuenta de cargo BCP debe tener 13 dígitos.");
        else if (Telecredit.ChargeAccountCurrency(account) is not string accountCurrency)
            ModelState.AddModelError(nameof(input.AccountNumber), "Revisa el dígito 11 de la cuenta BCP. Debe ser 0 para soles o 1 para dólares.");
        else if (input.Currency is "PEN" or "USD" && accountCurrency != input.Currency)
            ModelState.AddModelError(nameof(input.Currency), accountCurrency == "USD"
                ? "El dígito 11 es 1. Selecciona Dólares o ingresa el número de la cuenta en soles."
                : "El dígito 11 es 0. Selecciona Soles o ingresa el número de la cuenta en dólares.");
        if (!ModelState.IsValid) return View("Index", input);

        try
        {
            await repository.SaveChargeAccountAsync(Actor, new(input.AccountType, account, input.Currency, version), ct);
            TempData["TelecreditSaved"] = "Cuenta de cargo guardada.";
            return RedirectToAction(nameof(Index));
        }
        catch (SqlException error) when (error.Number == 50000)
        {
            ModelState.AddModelError("", "Vuelve a abrir la configuración de Telecrédito y revisa los datos.");
            return View("Index", input);
        }
    }
}

public sealed class TelecreditSettingsForm
{
    [Required, RegularExpression("C|M")] public string AccountType { get; set; } = "C";
    [Required, StringLength(30)] public string? AccountNumber { get; set; }
    [Required, RegularExpression("PEN|USD")] public string Currency { get; set; } = "PEN";
    public string? Version { get; set; }
    public bool CanEdit { get; set; }
}
