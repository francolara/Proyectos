using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using GestionServicios.Application.Payroll;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Web.Controllers;

[Authorize(Policy = "perm:TRABAJADORES.CONSULTAR")]
[Route("Planillas/Conceptos")]
public sealed class PayrollConceptsController(IPayrollConceptRepository repository) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit => User.HasClaim(AccessService.PermissionClaim, "TRABAJADORES.EDITAR");

    [HttpGet("")]
    public async Task<IActionResult> Index(string? search, string? nature, bool? active, int page = 1, CancellationToken ct = default)
    {
        if (!ModelState.IsValid || page < 1 || search?.Length > 100 || nature is not (null or "CARGO" or "ABONO"))
            return BadRequest();

        var concepts = await repository.ConceptsAsync(Actor, ct);
        var matches = concepts.Where(concept => (!active.HasValue || concept.Activo == active)
            && (nature is null || concept.Naturaleza == nature)
            && (string.IsNullOrWhiteSpace(search) || concept.Nombre.Contains(search.Trim(), StringComparison.CurrentCultureIgnoreCase))).ToArray();
        page = Math.Min(page, Math.Max(1, (matches.Length + 19) / 20));
        return View(new PayrollConceptsPage(matches.Skip((page - 1) * 20).Take(20).ToArray(),
            CanEdit, search, nature, active, page, matches.Length));
    }

    [HttpGet("Ficha/{id:long?}")]
    public async Task<IActionResult> Details(long? id, CancellationToken ct)
    {
        if (id is null) return CanEdit ? View(new PayrollConceptForm()) : Forbid();
        var concept = (await repository.ConceptsAsync(Actor, ct)).FirstOrDefault(item => item.ConceptoId == id);
        return concept is null ? NotFound() : View(new PayrollConceptForm
        {
            Id = concept.ConceptoId, Name = concept.Nombre, Nature = concept.Naturaleza, Calculation = concept.Calculo,
            Value = concept.Valor, Active = concept.Activo, Version = Convert.ToBase64String(concept.VersionFila),
            Pension = concept.Regimen is not null, ReadOnly = !CanEdit
        });
    }

    [HttpPost("Ficha")]
    [Authorize(Policy = "perm:TRABAJADORES.EDITAR")]
    public async Task<IActionResult> Save(PayrollConceptForm input, CancellationToken ct)
    {
        byte[] version = [];
        if (input.Id < 0) return BadRequest();
        if (input.Id > 0)
        {
            try { version = Convert.FromBase64String(input.Version ?? ""); }
            catch (FormatException) { return BadRequest(); }
            if (version.Length != 8) return BadRequest();
            var current = (await repository.ConceptsAsync(Actor, ct)).FirstOrDefault(item => item.ConceptoId == input.Id);
            if (current is null) return NotFound();
            input.Pension = current.Regimen is not null;
        }
        input.ReadOnly = false;
        if (input.Pension && input.Nature != "CARGO")
            ModelState.AddModelError(nameof(input.Nature), "El régimen previsional corresponde a un cargo.");
        if (input.Value is decimal value &&
            (value is < 0 or > 9999999999999999.99m ||
             (input.Calculation == "PORCENTAJE" ? value > 100 || decimal.Round(value, 4) != value : decimal.Round(value, 2) != value)))
            ModelState.AddModelError(nameof(input.Value), "Usa hasta dos decimales para un importe o un porcentaje entre 0 y 100 con hasta cuatro decimales.");

        if (ModelState.IsValid)
        {
            try
            {
                var id = await repository.SaveConceptAsync(Actor,
                    new(input.Id, input.Name.Trim(), input.Nature, input.Calculation, input.Value, null, input.Active, version), ct);
                TempData["PayrollConceptSaved"] = "Concepto guardado.";
                return RedirectToAction(nameof(Details), new { id });
            }
            catch (SqlException error) when (error.Number is 2601 or 2627)
            { ModelState.AddModelError(nameof(input.Name), "Ya existe un concepto con ese nombre."); }
            catch (SqlException error) when (error.Number == 50000)
            { ModelState.AddModelError("", "Vuelve a abrir la ficha y revisa los datos del concepto."); }
        }

        return View("Details", input);
    }
}

public sealed record PayrollConceptsPage(IReadOnlyList<PayrollConcept> Concepts, bool CanEdit, string? Search,
    string? Nature, bool? Active, int Page, int Total);

public sealed class PayrollConceptForm
{
    public long Id { get; set; }
    public string? Version { get; set; }
    [Required, StringLength(100)] public string Name { get; set; } = "";
    [Required, RegularExpression("CARGO|ABONO")] public string Nature { get; set; } = "ABONO";
    [Required, RegularExpression("IMPORTE|PORCENTAJE")] public string Calculation { get; set; } = "IMPORTE";
    public decimal? Value { get; set; }
    public bool Active { get; set; } = true;
    public bool Pension { get; set; }
    public bool ReadOnly { get; set; }
}
