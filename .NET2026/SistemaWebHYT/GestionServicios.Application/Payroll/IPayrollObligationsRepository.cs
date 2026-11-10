namespace GestionServicios.Application.Payroll;

public sealed record PayrollObligation(long ObligacionId, string Tipo, string Entidad, DateOnly Mes,
    decimal Importe, DateOnly Vencimiento, DateOnly? FechaPago, string? ReferenciaPago, byte[] VersionFila, string Estado);
public sealed record SctrPolicy(long PolizaId, string Aseguradora, string NumeroPoliza, string Cobertura,
    DateOnly VigenciaDesde, DateOnly VigenciaHasta, byte[] VersionFila, string Estado, SctrWorker[]? Trabajadores = null);
public sealed record SctrWorker(long TrabajadorId, string Nombre);
public sealed record PayrollObligations(DateOnly Mes, PayrollObligation[] Obligaciones, SctrPolicy[] Polizas, SctrWorker[]? PersonalActivo = null);

public sealed record PensionEstimate(string Entidad, string Regimen, decimal Porcentaje, decimal BaseMensual, decimal Estimado,
    PensionEstimateWorker[]? Trabajadores = null);
public sealed record PensionEstimateWorker(string Documento, string Nombre, decimal SueldoMensual, decimal Estimado);

public sealed record PayrollCatalogEntry(string Nombre, string? Regimen);
public sealed record PayrollCatalog(PayrollCatalogEntry[] Bancos, PayrollCatalogEntry[] Entidades);

public interface IPayrollCatalogRepository
{
    Task<PayrollCatalog> CatalogAsync(string actor, CancellationToken ct);
}

public interface IPayrollObligationsRepository
{
    Task<PayrollObligations> ObligationsAsync(string actor, DateOnly month, CancellationToken ct);
    Task<IReadOnlyList<PensionEstimate>> PensionEstimatesAsync(string actor, DateOnly month, CancellationToken ct, bool includeWorkers = false);
    Task<long> SavePensionRateAsync(string actor, string entity, decimal rate, CancellationToken ct);
    Task<long> SaveObligationAsync(string actor, PayrollObligation input, CancellationToken ct, bool preservePayment = false);
    Task<long> SavePolicyAsync(string actor, SctrPolicy input, CancellationToken ct);
}
