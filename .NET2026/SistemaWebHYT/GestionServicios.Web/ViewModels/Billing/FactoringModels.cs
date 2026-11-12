using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using GestionServicios.Application.Billing;

namespace GestionServicios.Web.ViewModels.Billing;

public sealed record FactoringSearch(DateOnly From, DateOnly To, string? Search, int Page = 1);
public sealed record FactoringList(FactoringSearch Filter, JsonElement Data, bool CanEdit);
public sealed record FactoringForm(FactoringSearch Filter, JsonElement Data, FactoringInput Input);

public sealed class FactoringInput
{
    [Required(ErrorMessage = "Selecciona o escribe la entidad financiera.")]
    [StringLength(150)]
    public string Financiera { get; set; } = "";
    public DateOnly FechaNegociacion { get; set; }
    public DateOnly? FechaDesembolso { get; set; }
    [Range(typeof(decimal), "0.0001", "100", ErrorMessage = "Ingresa un porcentaje mayor que cero y hasta 100.")]
    public decimal Porcentaje { get; set; }
    [StringLength(100)]
    public string? Referencia { get; set; }
    [StringLength(500)]
    public string? Observaciones { get; set; }
    public Guid OperacionId { get; set; } = Guid.NewGuid();
    public List<long> Facturas { get; set; } = [];

    public FactoringLot ToLot() => new(Financiera.Trim(), FechaNegociacion, FechaDesembolso, Porcentaje,
        Referencia?.Trim(), Observaciones?.Trim(), OperacionId, Facturas.Distinct().Order().ToArray());
}
