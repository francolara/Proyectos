using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:TRABAJADORES.CONSULTAR")]
[Authorize(Policy = "perm:ASISTENCIA.CONSULTAR")]
[Route("Planillas/Trabajadores/{id:long}/Asistencia")]
public sealed class AttendanceController(IWorkersRepository repository, TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "ASISTENCIA.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(long id, DateOnly? from, DateOnly? to, long? shiftId,
        int page = 1, CancellationToken ct = default)
    {
        if (!ModelState.IsValid || id <= 0 || page < 1 || from > to) return BadRequest();
        var today = TimeZoneInfo.ConvertTime(clock.GetUtcNow(), TimeZoneInfo.FindSystemTimeZoneById("America/Lima"));
        from ??= new DateOnly(today.Year, today.Month, 1);
        to ??= from.Value.AddDays(DateTime.DaysInMonth(from.Value.Year, from.Value.Month) - from.Value.Day);
        if (from > to) return BadRequest();

        var workers = await repository.WorkersAsync(Actor, id, null, null, 1, ct);
        if (workers.Count == 0) return NotFound();

        var shifts = await repository.ShiftsAsync(Actor, id, from, to, null, page, ct);
        WorkShift? selected = null;
        if (shiftId.HasValue)
        {
            var matches = await repository.ShiftsAsync(Actor, id, null, null, shiftId, 1, ct);
            if (matches.Count == 0) return NotFound();
            selected = matches[0];
        }
        return View(new AttendancePage(workers[0], shifts, selected, from.Value, to.Value, page, CanEdit));
    }

    [HttpPost("")]
    [Authorize(Policy = "perm:ASISTENCIA.EDITAR")]
    public async Task<IActionResult> Save(long id, AttendanceForm input, DateOnly? from, DateOnly? to, CancellationToken ct)
    {
        if (id <= 0) return BadRequest();
        byte[]? version = null;
        if (input.ShiftId.HasValue)
        {
            try { version = Convert.FromBase64String(input.Version ?? ""); }
            catch (FormatException) { return BadRequest(); }
            if (version.Length != 8) return BadRequest();
        }

        if (!ModelState.IsValid)
        {
            TempData["AttendanceMessage"] = "Revisa la entrada, la salida y los minutos extra.";
            return RedirectToAction(nameof(Index), new { id, from, to, shiftId = input.ShiftId });
        }

        try
        {
            await repository.SaveShiftAsync(Actor, id,
                new(input.ShiftId, version, input.Entry!.Value, input.Exit, input.ExtraMinutes), ct);
            TempData["AttendanceMessage"] = "Jornada guardada.";
            return RedirectToAction(nameof(Index), new { id, from, to });
        }
        catch (SqlException error) when (error.Number is 50000 or 2601 or 2627)
        {
            TempData["AttendanceMessage"] = error.Message == "La jornada se superpone con otra jornada del trabajador."
                ? "El horario coincide con otra jornada del trabajador."
                : "No se pudo guardar. Revisa el horario, los minutos extra y el estado de la ficha.";
            return RedirectToAction(nameof(Index), new { id, from, to, shiftId = input.ShiftId });
        }
    }
}

public sealed record AttendancePage(Worker Worker, IReadOnlyList<WorkShift> Shifts,
    WorkShift? Selected, DateOnly From, DateOnly To, int Page, bool CanEdit);

public sealed class AttendanceForm
{
    public long? ShiftId { get; set; }
    public string? Version { get; set; }
    [Required] public DateTime? Entry { get; set; }
    public DateTime? Exit { get; set; }
    [Range(0, int.MaxValue)] public int ExtraMinutes { get; set; }
}
