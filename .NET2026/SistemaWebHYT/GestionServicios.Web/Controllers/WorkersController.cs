using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:TRABAJADORES.CONSULTAR")]
[Route("Planillas/Trabajadores")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class WorkersController(IWorkersRepository repository) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "TRABAJADORES.EDITAR");
    private bool CanSeeAccount => User.HasClaim(AccessService.PermissionClaim, "TRABAJADORES.CUENTA");

    [HttpGet("")]
    public async Task<IActionResult> Index(string? search, string? kind, bool? active, int page = 1, CancellationToken ct = default)
    {
        if (!ModelState.IsValid || page < 1 || search?.Length > 180 || kind is not (null or "PLANILLA" or "EXTERNO")) return BadRequest();

        var workers = await repository.WorkersAsync(Actor, null, search, kind, page, ct, active);
        return View(new WorkersList(workers, search, kind, page, CanEdit, active));
    }

    [HttpGet("Ficha/{id:long?}")]
    public async Task<IActionResult> Details(long? id, CancellationToken ct)
    {
        if (id is null)
            return CanEdit ? await ShowAsync(new WorkerForm { Active = true }, ct) : Forbid();

        if (id <= 0) return NotFound();
        var records = await repository.WorkersAsync(Actor, id, null, null, 1, ct);
        if (records.Count == 0) return NotFound();

        var worker = records[0];
        return await ShowAsync(new WorkerForm
        {
            Id = worker.Id, Document = worker.Document, Name = worker.Name, Active = worker.Active,
            Kind = worker.Kind, StartDate = DateOnly.FromDateTime(worker.StartDate),
            Bank = worker.Bank, Account = CanSeeAccount ? worker.Account : worker.MaskedAccount,
            Cci = CanSeeAccount ? worker.Cci : null,
            Version = Convert.ToBase64String(worker.Version),
            CargoId = worker.CargoId, CargoName = worker.CargoName, MonthlySalary = worker.MonthlySalary,
            Entry = worker.PaySettings?.Entry, Exit = worker.PaySettings?.Exit, Pension = worker.PaySettings?.Pension,
            Afp = worker.PaySettings?.Afp, Frequency = worker.PaySettings?.Frequency
        }, ct);
    }

    [HttpPost("Ficha/{id:long?}")]
    [Authorize(Policy = "perm:TRABAJADORES.EDITAR")]
    public async Task<IActionResult> Save(long? id, WorkerForm input, CancellationToken ct)
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

        if (input.MonthlySalary is decimal salary && salary != decimal.Round(salary, 2))
            ModelState.AddModelError(nameof(input.MonthlySalary), "Usa hasta dos decimales para el sueldo.");

        if (input.Entry.HasValue != input.Exit.HasValue || input.Entry.HasValue && input.Entry == input.Exit)
            ModelState.AddModelError("", "Indica una hora habitual de entrada y otra de salida.");
        if (!WorkerForm.IsHalfHour(input.Entry) || !WorkerForm.IsHalfHour(input.Exit))
        {
            var previous = id.HasValue
                ? (await repository.WorkersAsync(Actor, id, null, null, 1, ct)).FirstOrDefault()?.PaySettings
                : null;
            if (!WorkerForm.IsHalfHour(input.Entry) && input.Entry != previous?.Entry)
                ModelState.AddModelError(nameof(input.Entry), "Selecciona una entrada en intervalos de media hora.");
            if (!WorkerForm.IsHalfHour(input.Exit) && input.Exit != previous?.Exit)
                ModelState.AddModelError(nameof(input.Exit), "Selecciona una salida en intervalos de media hora.");
        }
        if (input.Pension != "AFP" && !string.IsNullOrWhiteSpace(input.Afp))
            ModelState.AddModelError(nameof(input.Afp), "La administradora corresponde al régimen AFP.");

        if (CanSeeAccount && !string.IsNullOrWhiteSpace(input.Cci) && string.IsNullOrWhiteSpace(input.Bank))
            ModelState.AddModelError(nameof(input.Cci), "Selecciona el banco correspondiente al CCI.");

        if (ModelState.IsValid)
        {
            try
            {
                var savedId = await repository.SaveWorkerAsync(Actor,
                    new(id, input.Document, input.Name, input.Kind, input.StartDate!.Value, input.Active, version, CanSeeAccount,
                        CanSeeAccount ? input.Bank : null, CanSeeAccount ? input.Account : null,
                        input.CargoId, input.MonthlySalary, true,
                        new(input.Entry, input.Exit, input.Pension, input.Afp, input.Frequency), CanSeeAccount ? input.Cci : null, CanSeeAccount), ct);
                TempData["WorkerSaved"] = "Trabajador guardado.";
                return RedirectToAction(nameof(Details), new { id = savedId });
            }
            catch (SqlException error) when (error.Number is 2601 or 2627)
            {
                ModelState.AddModelError(nameof(input.Document), "Ya existe un trabajador con este documento.");
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

        return await ShowAsync(input, ct);
    }

    private async Task<ViewResult> ShowAsync(WorkerForm input, CancellationToken ct)
    {
        IReadOnlyList<WorkerPosition> positions = CanEdit ? await repository.PositionsAsync(Actor, ct) : [];
        return View("Details", new WorkerDetails(input, CanEdit, CanSeeAccount, positions));
    }
}

public sealed record WorkersList(IReadOnlyList<Worker> Workers, string? Search, string? Kind, int Page, bool CanEdit, bool? Active);
public sealed record WorkerDetails(WorkerForm Input, bool CanEdit, bool CanSeeAccount, IReadOnlyList<WorkerPosition> Positions);

public sealed class WorkerForm
{
    public static bool IsHalfHour(TimeOnly? time) => time is null || time.Value.Ticks % TimeSpan.FromMinutes(30).Ticks == 0;
    public long? Id { get; set; }
    public string? Version { get; set; }

    [Required(ErrorMessage = "Indica el documento."), StringLength(20)]
    public string Document { get; set; } = "";

    [Required(ErrorMessage = "Indica el nombre completo."), StringLength(180)]
    public string Name { get; set; } = "";

    [StringLength(80)]
    public string? Bank { get; set; }

    [StringLength(34)]
    public string? Account { get; set; }

    [RegularExpression("[0-9]{20}", ErrorMessage = "El CCI debe tener 20 dígitos.")]
    public string? Cci { get; set; }

    [Required, RegularExpression("PLANILLA|EXTERNO", ErrorMessage = "Selecciona un tipo de trabajador.")]
    public string Kind { get; set; } = "PLANILLA";

    [Required(ErrorMessage = "Indica la fecha de ingreso.")]
    public DateOnly? StartDate { get; set; }

    public long? CargoId { get; set; }
    public string? CargoName { get; set; }
    [Range(typeof(decimal), "0.01", "9999999999999999.99", ErrorMessage = "Indica un sueldo positivo.")]
    public decimal? MonthlySalary { get; set; }

    public TimeOnly? Entry { get; set; }
    public TimeOnly? Exit { get; set; }
    [RegularExpression("AFP|ONP")] public string? Pension { get; set; }
    [StringLength(100)] public string? Afp { get; set; }
    [RegularExpression("SEMANAL|QUINCENAL|MENSUAL")] public string? Frequency { get; set; }
    public bool Active { get; set; }
}
