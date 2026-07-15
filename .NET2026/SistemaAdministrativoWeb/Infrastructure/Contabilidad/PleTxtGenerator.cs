using System.Globalization;
using System.Text;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

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
            item.TipoComprobante,
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
            FormatearFecha(item.FechaOperacion),
            Sanitizar(item.Glosa),
            item.CodigoCuentaContable,
            item.CodigoMoneda,
            FormatearImporte(item.Debe),
            FormatearImporte(item.Haber),
            item.EstadoOperacion)), cancellationToken);
    }

    public Task<byte[]> GenerarLibroMayor61Async(IReadOnlyCollection<LibroMayor61Dto> items, CancellationToken cancellationToken = default)
    {
        return GenerarAsync(items.Select(item => string.Join('|',
            item.PeriodoPle,
            item.Cuo,
            item.CorrelativoMovimiento,
            item.CodigoCuentaContable,
            FormatearFecha(item.FechaOperacion),
            Sanitizar(item.Glosa),
            item.CodigoMoneda,
            FormatearImporte(item.Debe),
            FormatearImporte(item.Haber),
            item.EstadoOperacion)), cancellationToken);
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
        var codigoLibro = Sanitizar(item.CodigoLibroRelacionado);
        if (string.IsNullOrWhiteSpace(codigoLibro))
        {
            return Sanitizar(item.InformacionComplementaria);
        }

        return string.Concat(
            codigoLibro,
            Sanitizar(item.PeriodoPle),
            Sanitizar(item.Cuo),
            Sanitizar(item.CorrelativoMovimiento));
    }

}
