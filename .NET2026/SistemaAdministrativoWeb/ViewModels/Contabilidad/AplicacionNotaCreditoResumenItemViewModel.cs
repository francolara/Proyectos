namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AplicacionNotaCreditoResumenItemViewModel
{
    public int IdAplicacionNotaCredito { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public string TipoPersonaTexto { get; init; } = string.Empty;
    public string NombrePersona { get; init; } = string.Empty;
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public DateOnly FechaAplicacion { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal ImporteAplicado { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public string TipoComprobanteAplicado { get; init; } = string.Empty;
    public string DescripcionTipoComprobanteAplicado { get; init; } = string.Empty;
    public string SerieAplicado { get; init; } = string.Empty;
    public string NumeroAplicado { get; init; } = string.Empty;
    public string TipoComprobanteNc { get; init; } = string.Empty;
    public string DescripcionTipoComprobanteNc { get; init; } = string.Empty;
    public string SerieNc { get; init; } = string.Empty;
    public string NumeroNc { get; init; } = string.Empty;
    public string DocumentoAplicadoTexto => $"{TipoComprobanteAplicado} - {DescripcionTipoComprobanteAplicado} {SerieAplicado}-{NumeroAplicado}";
    public string DocumentoNcTexto => $"{TipoComprobanteNc} - {DescripcionTipoComprobanteNc} {SerieNc}-{NumeroNc}";
}
