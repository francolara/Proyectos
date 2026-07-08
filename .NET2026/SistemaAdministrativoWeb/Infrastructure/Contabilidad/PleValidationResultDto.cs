namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleValidationResultDto
{
    public IReadOnlyCollection<PleValidationIssueDto> Observaciones { get; init; } = [];
    public int CantidadErrores => Observaciones.Count(x => x.Severidad == PleValidationSeverity.Error);
    public int CantidadAdvertencias => Observaciones.Count(x => x.Severidad == PleValidationSeverity.Advertencia);
    public int CantidadInformacion => Observaciones.Count(x => x.Severidad == PleValidationSeverity.Informacion);
    public bool TieneErroresCriticos => CantidadErrores > 0;
}
