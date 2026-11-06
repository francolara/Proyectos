using System.Globalization;
using System.Security.Claims;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Maintenance;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:CONFIGURACION.CONSULTAR")]
[Route("Configuracion/TasasIGV")]
public sealed class TaxRateController(ITaxRateRepository repository, ILogger<TaxRateController> logger) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "CONFIGURACION.EDITAR");
    [HttpGet("")]
    public async Task<IActionResult> Index(string? search, int page = 1, CancellationToken ct = default)
    {
        if (search?.Length > 200 || page is < 1 or > 100000000) return BadRequest();
        try
        {
            var data = await repository.ListAsync(Actor, search, page, ct);
            return View(new TaxRateModel(data, new() { ExpectedLatest = data.Latest?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) }, CanEdit, search ?? "", page));
        }
        catch (Exception error) when (error is not OperationCanceledException)
        {
            logger.LogError(error, "No se pudo consultar vigencias IGV"); Response.StatusCode = 503;
            return View(new TaxRateModel(new([], 0, null), new(), false, search ?? "", page, "No se pudo cargar las vigencias. Vuelve a intentarlo."));
        }
    }

    [HttpPost("Agregar")]
    [Authorize(Policy = "perm:CONFIGURACION.EDITAR")]
    public async Task<IActionResult> Add(TaxRateInput input, CancellationToken ct)
    {
        var dateValid = DateOnly.TryParseExact(input.From, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var from);
        var valueValid = decimal.TryParse(input.Value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var value) && value is >= 0 and <= 1 && decimal.Round(value, 6) == value;
        if (!dateValid) ModelState.AddModelError("From", "Selecciona una fecha de inicio válida.");
        if (!valueValid) ModelState.AddModelError("Value", "Ingresa una tasa entre 0 y 1, con hasta seis decimales; por ejemplo, 0.18.");
        if (string.IsNullOrWhiteSpace(input.Reason) || input.Reason.Length > 1000) ModelState.AddModelError("Reason", "Indica un motivo de hasta 1000 caracteres.");
        DateOnly? expected = null;
        if (!string.IsNullOrWhiteSpace(input.ExpectedLatest))
        {
            if (DateOnly.TryParseExact(input.ExpectedLatest, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var latest)) expected = latest;
            else ModelState.AddModelError("", "Vuelve a cargar las vigencias antes de continuar.");
        }
        if (ModelState.IsValid)
        {
            try
            {
                await repository.AddAsync(Actor, from, value, input.Reason!.Trim(), expected, ct);
                TempData["Notice"] = "Nueva vigencia registrada. Se conserva el histórico de presupuestos.";
                return RedirectToAction(nameof(Index));
            }
            catch (MaintenanceException error) { ModelState.AddModelError("", error.Message); }
            catch (Exception error) when (error is not OperationCanceledException)
            { logger.LogError(error, "No se pudo registrar vigencia IGV"); ModelState.AddModelError("", "No se pudo confirmar el registro. Consulta las vigencias antes de reintentar."); }
        }
        var data = await repository.ListAsync(Actor, null, 1, ct);
        return View("Index", new TaxRateModel(data, input, CanEdit, "", 1));
    }
}
