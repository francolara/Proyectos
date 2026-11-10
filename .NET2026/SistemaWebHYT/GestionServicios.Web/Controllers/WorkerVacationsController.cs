using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:TRABAJADORES.CONSULTAR")]
[Authorize(Policy = "perm:VACACIONES.CONSULTAR")]
[Route("Planillas/Trabajadores/{id:long}/Vacaciones")]
public sealed class WorkerVacationsController(IWorkersRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;

    [HttpGet("")]
    public async Task<IActionResult> Index(long id, long? entryId, CancellationToken ct)
    {
        var workers = await repository.WorkersAsync(Actor, id, null, null, 1, ct);
        if (workers.Count == 0) return NotFound();

        var today = TimeZoneInfo.ConvertTime(clock.GetUtcNow(), TimeZoneInfo.FindSystemTimeZoneById("America/Lima"));
        var data = await repository.VacationsAsync(Actor, id, DateOnly.FromDateTime(today.DateTime), ct);
        var entry = entryId is null ? null : data.Entries.FirstOrDefault(item => item.Id == entryId);
        if (entryId is not null && entry is null) return NotFound();

        var worker = workers[0];
        var canEdit = worker.Kind == "PLANILLA" && User.HasClaim(AccessService.PermissionClaim, "VACACIONES.EDITAR");
        return View(new WorkerVacationsPage(id, worker.Name, data, entry, canEdit));
    }

    [HttpPost("Periodos")]
    [Authorize(Policy = "perm:VACACIONES.EDITAR")]
    public async Task<IActionResult> AddPeriod(long id, VacationDatesForm input, CancellationToken ct)
    {
        if (!ModelState.IsValid || input.To < input.From)
            return Finish(id, "Revisa las fechas y los días reconocidos.");

        try
        {
            await repository.SaveVacationPeriodAsync(Actor, id, new(input.From, input.To, input.Days), ct);
            return Finish(id, "Período registrado.");
        }
        catch (SqlException error) when (error.Number == 50000)
        {
            return Finish(id, "No se pudo registrar el período. Revisa las fechas y los períodos existentes.");
        }
    }

    [HttpPost("")]
    [Authorize(Policy = "perm:VACACIONES.EDITAR")]
    public async Task<IActionResult> Save(long id, VacationEntryForm input, CancellationToken ct)
    {
        byte[]? version = null;
        if (input.EntryId.HasValue)
        {
            try { version = Convert.FromBase64String(input.Version ?? ""); }
            catch (FormatException) { return BadRequest(); }
            if (version.Length != 8) return BadRequest();
        }

        if (!ModelState.IsValid || input.To < input.From)
            return Finish(id, "Revisa las fechas, los días y el período.", input.EntryId);

        try
        {
            await repository.SaveVacationAsync(Actor, id,
                new(input.PeriodId, input.From, input.To, input.Days, input.Status, input.EntryId, version), ct);
            return Finish(id, "Vacaciones guardadas.");
        }
        catch (SqlException error) when (error.Number == 50000)
        {
            var message = error.Message switch
            {
                "Los dias exceden el saldo disponible del periodo." => "Los días exceden el saldo disponible del período.",
                "Las fechas se superponen con otras vacaciones registradas." => "Las fechas coinciden con otras vacaciones registradas.",
                _ => "No se pudo guardar el registro. Revisa sus datos y vuelve a intentarlo."
            };
            return Finish(id, message, input.EntryId);
        }
    }

    private RedirectToActionResult Finish(long id, string message, long? entryId = null)
    {
        TempData["VacationMessage"] = message;
        return RedirectToAction(nameof(Index), new { id, entryId });
    }
}

public sealed record WorkerVacationsPage(long WorkerId, string WorkerName, WorkerVacations Data,
    VacationEntry? Entry, bool CanEdit);

public class VacationDatesForm
{
    [Required] public DateOnly From { get; set; }
    [Required] public DateOnly To { get; set; }
    [Range(1, short.MaxValue)] public short Days { get; set; }
}

public sealed class VacationEntryForm : VacationDatesForm
{
    [Range(typeof(long), "1", "9223372036854775807")] public long PeriodId { get; set; }
    [Required, RegularExpression("PENDIENTE|APROBADA|ANULADA")] public string Status { get; set; } = "";
    public long? EntryId { get; set; }
    public string? Version { get; set; }
}
