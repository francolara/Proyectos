using System.Globalization;
using System.Text;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

// Firma: FRANCO LARA - 03/08/2026 | Exporta los 21 campos base de 5.1, 5.2 y 6.1, conserva el palote final y usa 00 cuando falta el tipo de comprobante.
public sealed class PleTxtGenerator : IPleTxtGenerator
{
    private static readonly UTF8Encoding Utf8SinBom = new(false);

    public Task<byte[]> GenerarLibroDiario51Async(IReadOnlyCollection<LibroDiario51Dto> items, CancellationToken cancellationToken = default)
    {
        return GenerarAsync(items.Select(item => string.Join('|',
            item.PeriodoPle,
            item.Cuo,
            item.CorrelativoMovimiento,
            item.CodigoCuentaContable,
            item.CodigoUnidadOperacion,
            item.CodigoCentroCosto,
            item.CodigoMoneda,
            item.TipoDocumentoEmisor,
            item.NumeroDocumentoEmisor,
            FormatearTipoComprobante(item.TipoComprobante),
            item.SerieComprobante,
            item.NumeroComprobante,
            FormatearFecha(item.FechaContable),
            FormatearFecha(item.FechaVencimiento),
            FormatearFecha(item.FechaOperacion),
            Sanitizar(item.Glosa),
            Sanitizar(item.GlosaReferencial),
            FormatearImporte(item.Debe),
            FormatearImporte(item.Haber),
            FormatearReferenciaEstructuradaLibroDiario51(item),
            item.EstadoOperacion) + '|'), cancellationToken);
    }

    public Task<byte[]> GenerarLibroDiario52Async(IReadOnlyCollection<LibroDiario52Dto> items, CancellationToken cancellationToken = default)
    {
        return GenerarAsync(items.Select(item => string.Join('|',
            item.PeriodoPle,
            item.Cuo,
            item.CorrelativoAsiento,
            item.CodigoCuentaContable,
            item.CodigoUnidadOperacion,
            item.CodigoCentroCosto,
            item.CodigoMoneda,
            item.TipoDocumentoEmisor,
            item.NumeroDocumentoEmisor,
            FormatearTipoComprobante(item.TipoComprobante),
            item.SerieComprobante,
            item.NumeroComprobante,
            FormatearFecha(item.FechaContable),
            FormatearFecha(item.FechaVencimiento),
            FormatearFecha(item.FechaOperacion),
            Sanitizar(item.Glosa),
            Sanitizar(item.GlosaReferencial),
            FormatearImporte(item.Debe),
            FormatearImporte(item.Haber),
            FormatearReferenciaEstructurada(item.CodigoLibroRelacionado, item.PeriodoPle, item.Cuo, item.CorrelativoAsiento, item.InformacionComplementaria),
            item.EstadoOperacion) + '|'), cancellationToken);
    }

    public Task<byte[]> GenerarLibroMayor61Async(IReadOnlyCollection<LibroMayor61Dto> items, CancellationToken cancellationToken = default)
    {
        return GenerarAsync(items.Select(item => string.Join('|',
            item.PeriodoPle,
            item.Cuo,
            item.CorrelativoMovimiento,
            item.CodigoCuentaContable,
            item.CodigoUnidadOperacion,
            item.CodigoCentroCosto,
            item.CodigoMoneda,
            item.TipoDocumentoEmisor,
            item.NumeroDocumentoEmisor,
            FormatearTipoComprobante(item.TipoComprobante),
            item.SerieComprobante,
            item.NumeroComprobante,
            FormatearFecha(item.FechaContable),
            FormatearFecha(item.FechaVencimiento),
            FormatearFecha(item.FechaOperacion),
            Sanitizar(item.Glosa),
            Sanitizar(item.GlosaReferencial),
            FormatearImporte(item.Debe),
            FormatearImporte(item.Haber),
            FormatearReferenciaEstructurada(item.CodigoLibroRelacionado, item.PeriodoPle, item.Cuo, item.CorrelativoMovimiento, item.InformacionComplementaria),
            item.EstadoOperacion) + '|'), cancellationToken);
    }

    private static Task<byte[]> GenerarAsync(IEnumerable<string> lineas, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var contenido = string.Join(Environment.NewLine, lineas);
        return Task.FromResult(Utf8SinBom.GetBytes(contenido));
    }

    private static string FormatearImporte(decimal importe)
    {
        return Math.Round(importe, 2, MidpointRounding.AwayFromZero).ToString("0.00", CultureInfo.InvariantCulture);
    }

    private static string FormatearFecha(DateOnly? fecha)
    {
        return fecha.HasValue ? fecha.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : string.Empty;
    }

    private static string FormatearTipoComprobante(string? tipoComprobante)
    {
        return string.IsNullOrWhiteSpace(tipoComprobante) ? "00" : tipoComprobante.Trim();
    }

    private static string Sanitizar(string? texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
        {
            return string.Empty;
        }

        return texto.Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Replace("|", "/", StringComparison.Ordinal)
            .Trim();
    }

    private static string FormatearReferenciaEstructuradaLibroDiario51(LibroDiario51Dto item)
    {
        return FormatearReferenciaEstructurada(
            item.CodigoLibroRelacionado,
            item.PeriodoPle,
            item.Cuo,
            item.CorrelativoMovimiento,
            item.InformacionComplementaria);
    }

    private static string FormatearReferenciaEstructurada(string codigoLibroRelacionado, string periodoPle, string cuo, string correlativo, string informacionComplementaria)
    {
        var codigoLibro = Sanitizar(codigoLibroRelacionado);
        if (string.IsNullOrWhiteSpace(codigoLibro))
        {
            return Sanitizar(informacionComplementaria);
        }

        return string.Concat(
            codigoLibro,
            Sanitizar(periodoPle),
            Sanitizar(cuo),
            Sanitizar(correlativo));
    }

}
