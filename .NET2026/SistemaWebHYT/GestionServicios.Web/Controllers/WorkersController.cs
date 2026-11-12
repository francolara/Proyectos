using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Reporting;
using GestionServicios.Application.Maintenance;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 11/11/2026 | Imprime el listado completo de trabajadores con sus filtros y el número de cuenta enmascarado.
[Authorize(Policy = "perm:TRABAJADORES.CONSULTAR")]
[Route("Planillas/Trabajadores")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class WorkersController(IWorkersRepository repository, IPayrollConceptRepository concepts) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "TRABAJADORES.EDITAR");
    private bool CanSeeAccount => User.HasClaim(AccessService.PermissionClaim, "TRABAJADORES.CUENTA");

    [HttpGet("")]
    public async Task<IActionResult> Index(string? search, string? kind, bool? active, int page = 1, CancellationToken ct = default, bool print = false)
    {
        if (!ModelState.IsValid || page < 1 || search?.Length > 180 || kind is not (null or "PLANILLA" or "EXTERNO")) return BadRequest();

        Task<IReadOnlyList<Worker>> Load(int number) => repository.WorkersAsync(Actor, null, search, kind, number, ct, active);
        if (print)
        {
            if (!User.HasClaim(AccessService.PermissionClaim, "REPORTES.IMPRIMIR")) return Forbid();
            try
            {
                var rows = await ListingReport.AllAsync(async number =>
                {
                    var result = await Load(number);
                    return (result, result.FirstOrDefault()?.Total ?? 0);
                }, (Worker row) => row.Id);
                var summary = new[] { ListingReport.Filter("Buscar", search), ListingReport.Filter("Tipo",
                    kind == "PLANILLA" ? "En planilla" : kind == "EXTERNO" ? "Externo" : null),
                    ListingReport.Filter("Estado", active is bool status ? status ? "Activo" : "Inactivo" : null) };
                return View("ListingReport", new ListingReport("Trabajadores", summary, rows.Count,
                    new ListingTable(null, ["Documento", "Apellidos y nombres", "Cargo", "Tipo", "Ingreso", "Frecuencia", "Estado", "Banco", "Cuenta"],
                        rows.Select(row => new[] { row.Document, row.Name, row.CargoName ?? "", row.Kind == "PLANILLA" ? "En planilla" : "Externo",
                            ListingReport.Date(row.StartDate), row.PaySettings?.Frequency ?? "", row.Active ? "Activo" : "Inactivo",
                            row.Bank ?? "", row.MaskedAccount ?? "" }))));
            }
            catch (MaintenanceException error) { return Conflict(error.Message); }
        }
        var workers = await Load(page);
        return View(new WorkersList(workers, search, kind, page, CanEdit, active));
    }

    // Firma: FRANCO LARA - 11/11/2026 | Frecuencia quincenal inicial en el alta del trabajador.
    [HttpGet("Ficha/{id:long?}")]
    public async Task<IActionResult> Details(long? id, CancellationToken ct)
    {
        if (id is null)
            return CanEdit ? await ShowAsync(new WorkerForm { Active = true, Frequency = "QUINCENAL" }, ct) : Forbid();

        if (id <= 0) return NotFound();
        var records = await repository.WorkersAsync(Actor, id, null, null, 1, ct);
        if (records.Count == 0) return NotFound();

        var worker = records[0];
        return await ShowAsync(new WorkerForm
        {
            Id = worker.Id, Document = worker.Document, Name = worker.Name, Active = worker.Active,
            GivenNames = worker.GivenNames, Surnames = worker.Surnames,
            Kind = worker.Kind, StartDate = DateOnly.FromDateTime(worker.StartDate),
            Bank = worker.Bank, Account = CanSeeAccount ? worker.Account : worker.MaskedAccount,
            Cci = CanSeeAccount ? worker.Cci : null,
            BcpDocumentType = worker.Telecredit?.DocumentType,
            BcpAccountType = CanSeeAccount ? worker.Telecredit?.AccountType : null,
            BcpCurrency = CanSeeAccount ? worker.Telecredit?.Currency : null,
            Version = Convert.ToBase64String(worker.Version),
            CargoId = worker.CargoId, CargoName = worker.CargoName, MonthlySalary = worker.MonthlySalary,
            Entry = worker.PaySettings?.Entry, Exit = worker.PaySettings?.Exit, Pension = worker.PaySettings?.Pension,
            Afp = worker.PaySettings?.Afp, Frequency = worker.PaySettings?.Frequency, Concepts = worker.Concepts ?? []
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

        input.Name = $"{input.Surnames?.Trim()} {input.GivenNames?.Trim()}".Trim();
        if (input.Name.Length > 180)
            ModelState.AddModelError("", "Los nombres y apellidos juntos admiten hasta 180 caracteres.");

        var hasAccount = !string.IsNullOrWhiteSpace(input.Account) || !string.IsNullOrWhiteSpace(input.Cci);
        if (CanSeeAccount && string.IsNullOrWhiteSpace(input.Bank) == hasAccount)
            ModelState.AddModelError("", "Completa el banco y el número de cuenta o CCI.");

        var document = input.Document?.Trim() ?? "";
        if (input.EditTelecredit && input.BcpDocumentType is not null &&
            !Telecredit.ValidIdentity(input.BcpDocumentType, document))
            ModelState.AddModelError(nameof(input.Document), "El número no corresponde al tipo de documento seleccionado.");

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

        if (input.Concepts.Length > 200 || input.Concepts.Any(id => id <= 0))
            ModelState.AddModelError(nameof(input.Concepts), "Revisa los conceptos seleccionados.");

        if (ModelState.IsValid)
        {
            try
            {
                var savedId = await repository.SaveWorkerAsync(Actor,
                    new(id, document, input.Name, input.Kind, input.StartDate!.Value, input.Active, version, CanSeeAccount,
                        CanSeeAccount ? input.Bank : null, CanSeeAccount ? input.Account : null,
                        input.CargoId, input.MonthlySalary, true,
                        new(input.Entry, input.Exit, input.Pension, input.Afp, input.Frequency), CanSeeAccount ? input.Cci : null, CanSeeAccount,
                        input.EditConcepts ? input.Concepts : null,
                        input.EditTelecredit ? new(input.BcpDocumentType, null,
                            CanSeeAccount ? input.BcpAccountType : null, CanSeeAccount ? input.BcpCurrency : null) : null,
                        CanSeeAccount && input.EditTelecredit, input.GivenNames?.Trim(), input.Surnames?.Trim()), ct);
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
            input.BcpAccountType = null;
            input.BcpCurrency = null;
        }

        return await ShowAsync(input, ct);
    }

    private async Task<ViewResult> ShowAsync(WorkerForm input, CancellationToken ct)
    {
        IReadOnlyList<WorkerPosition> positions = CanEdit ? await repository.PositionsAsync(Actor, ct) : [];
        var options = await concepts.ConceptsAsync(Actor, ct);
        return View("Details", new WorkerDetails(input, CanEdit, CanSeeAccount, positions, options));
    }
}

public sealed record WorkersList(IReadOnlyList<Worker> Workers, string? Search, string? Kind, int Page, bool CanEdit, bool? Active);
public sealed record WorkerDetails(WorkerForm Input, bool CanEdit, bool CanSeeAccount, IReadOnlyList<WorkerPosition> Positions, IReadOnlyList<PayrollConcept> Concepts);

public sealed class WorkerForm
{
    public static bool IsHalfHour(TimeOnly? time) => time is null || time.Value.Ticks % TimeSpan.FromMinutes(30).Ticks == 0;
    public long? Id { get; set; }
    public string? Version { get; set; }

    [Required(ErrorMessage = "Indica el documento."), StringLength(20)]
    public string Document { get; set; } = "";

    [BindNever] public string Name { get; set; } = "";
    [Required(ErrorMessage = "Indica los nombres."), StringLength(179)]
    public string? GivenNames { get; set; }
    [Required(ErrorMessage = "Indica los apellidos."), StringLength(179)]
    public string? Surnames { get; set; }

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
    public bool EditTelecredit { get; set; }
    [RegularExpression("1|3|4")] public string? BcpDocumentType { get; set; }
    [RegularExpression("A|B|C|M")] public string? BcpAccountType { get; set; }
    [RegularExpression("PEN|USD")] public string? BcpCurrency { get; set; }
    public bool EditConcepts { get; set; }
    public long[] Concepts { get; set; } = [];
}
