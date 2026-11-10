using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:TRABAJADORES.CONSULTAR")]
[Route("Planillas/Cargos")]
public sealed class WorkerPositionsController(IWorkersRepository repository) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "TRABAJADORES.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(string? search, int page = 1, CancellationToken ct = default)
    {
        if (!ModelState.IsValid || page < 1 || search?.Length > 100) return BadRequest();

        var positions = await repository.PositionsAsync(Actor, ct);
        var matches = positions.Where(position => string.IsNullOrWhiteSpace(search)
            || position.Name.Contains(search.Trim(), StringComparison.CurrentCultureIgnoreCase)).ToArray();
        var maxPage = Math.Max(1, (matches.Length + 19) / 20);
        page = Math.Min(page, maxPage);
        return View(new WorkerPositionsPage(matches.Skip((page - 1) * 20).Take(20).ToArray(),
            CanEdit, search, page, matches.Length));
    }

    [HttpGet("Ficha/{id:long?}")]
    [Authorize(Policy = "perm:TRABAJADORES.EDITAR")]
    public async Task<IActionResult> Details(long? id, CancellationToken ct)
    {
        if (id is null) return View(new WorkerPositionForm());

        var positions = await repository.PositionsAsync(Actor, ct);
        var selected = positions.FirstOrDefault(position => position.Id == id);
        if (selected is null) return NotFound();
        return View(new WorkerPositionForm
        {
            Id = selected.Id, Name = selected.Name, Active = selected.Active,
            Version = Convert.ToBase64String(selected.Version)
        });
    }

    [HttpPost("Ficha")]
    [Authorize(Policy = "perm:TRABAJADORES.EDITAR")]
    public async Task<IActionResult> Save(WorkerPositionForm input, CancellationToken ct)
    {
        byte[] version = [];
        if (input.Id < 0) return BadRequest();
        if (input.Id > 0)
        {
            try { version = Convert.FromBase64String(input.Version ?? ""); }
            catch (FormatException) { return BadRequest(); }
            if (version.Length != 8) return BadRequest();
        }

        if (ModelState.IsValid)
        {
            try
            {
                await repository.SavePositionAsync(Actor, new(input.Id, input.Name, input.Active, version), ct);
                TempData["PositionSaved"] = "Cargo guardado.";
                return RedirectToAction(nameof(Index));
            }
            catch (SqlException error) when (error.Number is 2601 or 2627)
            {
                ModelState.AddModelError(nameof(input.Name), "Ya existe un cargo con ese nombre.");
            }
            catch (SqlException error) when (error.Number == 50000)
            {
                ModelState.AddModelError("", "Vuelve a abrir el cargo y revisa los datos.");
            }
        }

        return View("Details", input);
    }
}

public sealed record WorkerPositionsPage(IReadOnlyList<WorkerPosition> Positions,
    bool CanEdit, string? Search, int Page, int Total);

public sealed class WorkerPositionForm
{
    public long Id { get; set; }
    [Required(ErrorMessage = "Indica el nombre del cargo."), StringLength(100)]
    public string Name { get; set; } = "";
    public bool Active { get; set; } = true;
    public string? Version { get; set; }
}
