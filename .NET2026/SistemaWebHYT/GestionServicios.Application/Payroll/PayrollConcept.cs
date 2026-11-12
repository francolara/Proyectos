namespace GestionServicios.Application.Payroll;

public sealed record PayrollConcept(long ConceptoId, string Nombre, string Naturaleza, string Calculo,
    decimal? Valor, string? Regimen, bool Activo, byte[] VersionFila);

public sealed record PayPlanConcept(long ConceptoId, string Nombre, string Naturaleza, string Calculo, decimal? Valor, bool BaseMensual = false);

public interface IPayrollConceptRepository
{
    Task<IReadOnlyList<PayrollConcept>> ConceptsAsync(string actor, CancellationToken ct);
    Task<long> SaveConceptAsync(string actor, PayrollConcept concept, CancellationToken ct);
}
