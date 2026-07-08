namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPleValidationService
{
    Task<PleValidationResultDto> ValidarAsync(
        LibroElectronicoConsultaRequest request,
        string empresa,
        string ruc,
        IReadOnlyCollection<AsientoResumenDto> asientos,
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        IReadOnlyCollection<MonedaDto> monedas,
        IReadOnlyCollection<TipoDocumentoIdentidadDto> tiposDocumento,
        IReadOnlyCollection<TipoComprobanteDto> tiposComprobante,
        IReadOnlyCollection<LibroDiario51Dto> libroDiario51Items,
        IReadOnlyCollection<LibroDiario52Dto> libroDiario52Items,
        IReadOnlyCollection<LibroMayor61Dto> libroMayor61Items,
        CancellationToken cancellationToken = default);
}
