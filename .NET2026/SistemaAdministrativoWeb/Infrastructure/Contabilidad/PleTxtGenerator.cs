using System.Globalization;
using System.Text;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

// Firma: FRANCO LARA - 03/08/2026 | Exporta los 21 campos base de 5.1, 5.2 y 6.1, conserva el palote final y usa 00 cuando falta el tipo de comprobante.
// Firma: FRANCO LARA - 04/08/2026 | Genera referencias de Compras/Ventas y exporta planes 5.3/5.4 completos o incrementales con estados SUNAT 1 y 9.
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
            FormatearReferenciaEstructurada(item.CodigoLibroRelacionado, item.PeriodoReferencia, item.CuoReferencia, item.CorrelativoReferencia),
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
            FormatearReferenciaEstructurada(item.CodigoLibroRelacionado, item.PeriodoReferencia, item.CuoReferencia, item.CorrelativoReferencia),
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
            FormatearReferenciaEstructurada(item.CodigoLibroRelacionado, item.PeriodoReferencia, item.CuoReferencia, item.CorrelativoReferencia),
            item.EstadoOperacion) + '|'), cancellationToken);
    }

    public Task<byte[]> GenerarPlanContableAsync(IReadOnlyCollection<PlePlanCuentaExportItemDto> items, CancellationToken cancellationToken = default)
    {
        return GenerarAsync(items
            .OrderBy(item => item.CodigoCuenta, StringComparer.OrdinalIgnoreCase)
            .Select(item => string.Join('|',
                item.PeriodoPle,
                Sanitizar(item.CodigoCuenta),
                Limitar(Sanitizar(item.NombreCuenta), 100),
                "01",
                string.Empty,
                string.Empty,
                string.Empty,
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

    private static string Limitar(string texto, int longitudMaxima)
    {
        return texto.Length <= longitudMaxima ? texto : texto[..longitudMaxima];
    }

    private static string FormatearReferenciaEstructurada(string codigoLibroRelacionado, string periodoReferencia, string cuoReferencia, string correlativoReferencia)
    {
        var codigoLibro = Sanitizar(codigoLibroRelacionado);
        var periodo = Sanitizar(periodoReferencia);
        var cuo = Sanitizar(cuoReferencia);
        var correlativo = Sanitizar(correlativoReferencia);

        return string.IsNullOrWhiteSpace(codigoLibro)
            || string.IsNullOrWhiteSpace(periodo)
            || string.IsNullOrWhiteSpace(cuo)
            || string.IsNullOrWhiteSpace(correlativo)
                ? string.Empty
                : string.Join('&', codigoLibro, periodo, cuo, correlativo);
    }

}
