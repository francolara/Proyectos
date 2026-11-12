namespace GestionServicios.Application.Payroll;

public sealed record PayPlan(long? PlanId, byte[]? VersionFila, string Frecuencia, DateOnly Desde,
    DateOnly Hasta, PayPlanRow[] Filas, bool VerCuenta = false, DateOnly? FechaPago = null, string? ReferenciaPago = null, PayPlanConcept[]? Conceptos = null,
    bool PorIncorporar = false);

public sealed record PayPlanRow(long TrabajadorId, string Nombre, string Tipo, decimal Base,
    int MinutosExtra, decimal ImporteExtra, decimal Descuento, PayPlanAdjustment[]? Ajustes = null,
    string? Banco = null, string? Cuenta = null, string? Cci = null, decimal? SueldoMensual = null, string? Cargo = null, bool Incluido = true)
{
    public decimal Neto => Base + ImporteExtra - Descuento;
}

public sealed class PayPlanLineInput
{
    public long TrabajadorId { get; set; }
    public decimal Base { get; set; }
    public int MinutosExtra { get; set; }
    public decimal ImporteExtra { get; set; }
    public decimal Descuento { get; set; }
    public PayPlanAdjustment[]? Ajustes { get; set; }
    public bool AjustesDetallados { get; set; }
    public bool Incluido { get; set; } = true;
}

public sealed record PayPlanAdjustment(string Tipo, DateOnly? Fecha, string? Presupuesto, string? Concepto,
    int Horas, decimal Importe, long? ConceptoId = null, decimal? Porcentaje = null, decimal? BasePorcentaje = null);

public interface IPayPlanningRepository
{
    Task<PayPlan> PlanAsync(string actor, string frequency, DateOnly from, CancellationToken ct);
    Task<PayPlan> ReportPlanAsync(string actor, string frequency, DateOnly from, CancellationToken ct);
    Task<long> ConfirmPlanAsync(string actor, string frequency, DateOnly from, byte[] version,
        DateOnly date, string reference, CancellationToken ct);
    Task<long> SavePlanAsync(string actor, string frequency, DateOnly from, byte[]? version,
        IReadOnlyList<PayPlanLineInput> lines, CancellationToken ct);
}
