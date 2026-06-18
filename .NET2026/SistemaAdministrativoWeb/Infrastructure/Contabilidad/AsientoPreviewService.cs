namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoPreviewService(IConfiguracionContabilizacionRepository configuracionRepository) : IAsientoPreviewService
{
    public async Task<AsientoPreviewResultDto> PrevisualizarAsync(int idEmpresa, AsientoPreviewRequest request, CancellationToken cancellationToken = default)
    {
        ValidarMontos(request);
        ValidarDetalle(request.Detalles, request.BaseImponible);

        var moduloOperacion = (request.ModuloOperacion ?? string.Empty).Trim().ToUpperInvariant();
        var configuracion = await configuracionRepository.ObtenerAsync(request.IdConfiguracionContabilizacion, cancellationToken);

        if (configuracion is null || configuracion.IdEmpresa != idEmpresa || !string.Equals(configuracion.ModuloOperacion, moduloOperacion, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("La configuracion contable indicada no existe para la empresa activa.");
        }

        if (!configuracion.Activo)
        {
            throw new InvalidOperationException("La configuracion contable seleccionada esta inactiva.");
        }

        if (!configuracion.GeneraAsientoAutomatico)
        {
            throw new InvalidOperationException("La configuracion seleccionada no esta habilitada para generar asiento automatico.");
        }

        var prefijoGlosa = moduloOperacion == "VEN" ? "Venta" : "Compra";
        var referencia = ConstruirReferencia(request.TipoComprobante, request.Serie, request.Numero);
        var lineas = configuracion.Detalles
            .Where(x => x.Activo)
            .OrderBy(x => x.Orden)
            .Select((x, index) => new AsientoPreviewLineaDto
            {
                Item = (short)(index + 1),
                ComponenteContable = x.ComponenteContable,
                CodigoCuenta = x.CodigoCuenta,
                NombreCuenta = x.NombreCuenta,
                NaturalezaMovimiento = x.NaturalezaMovimiento,
                Debe = x.NaturalezaMovimiento == "D" ? ResolverMonto(x.ComponenteContable, request) : 0m,
                Haber = x.NaturalezaMovimiento == "H" ? ResolverMonto(x.ComponenteContable, request) : 0m,
                GlosaDetalle = $"{prefijoGlosa} {referencia} / {x.ComponenteContable}"
            })
            .Where(x => x.Debe != 0m || x.Haber != 0m)
            .ToList();

        if (lineas.Count == 0)
        {
            throw new InvalidOperationException($"La configuracion seleccionada no genera lineas contables con los importes de la {(moduloOperacion == "VEN" ? "venta" : "compra")}.");
        }

        var totalDebe = lineas.Sum(x => x.Debe);
        var totalHaber = lineas.Sum(x => x.Haber);
        var cuadrado = totalDebe == totalHaber;

        return new AsientoPreviewResultDto
        {
            ModuloOperacion = moduloOperacion,
            IdConfiguracionContabilizacion = configuracion.IdConfiguracionContabilizacion,
            EscenarioOperacion = configuracion.EscenarioOperacion,
            CodigoOrigen = configuracion.CodigoOrigen,
            NombreOrigen = configuracion.NombreOrigen,
            GlosaAsiento = $"{prefijoGlosa} {referencia}",
            TotalDebe = totalDebe,
            TotalHaber = totalHaber,
            Cuadrado = cuadrado,
            MensajeValidacion = cuadrado
                ? "El asiento propuesto esta cuadrado."
                : $"La configuracion no cuadra el asiento. Diferencia: {(totalDebe - totalHaber):N2}.",
            Lineas = lineas
        };
    }

    private static void ValidarMontos(AsientoPreviewRequest request)
    {
        if (request.BaseImponible < 0m
            || request.Igv < 0m
            || request.Isc < 0m
            || request.OtrosTributos < 0m
            || request.Redondeo < 0m
            || request.ImporteTotal < 0m)
        {
            throw new InvalidOperationException("Los montos no pueden ser negativos.");
        }

        if (request.ImporteTotal != request.BaseImponible + request.Igv + request.Isc + request.OtrosTributos + request.Redondeo)
        {
            throw new InvalidOperationException("El importe total debe ser igual a la suma de base imponible, IGV, ISC, otros tributos y redondeo.");
        }
    }

    private static void ValidarDetalle(IReadOnlyCollection<AsientoPreviewDetalleRequest> detalles, decimal baseImponible)
    {
        var detallesValidos = detalles
            .Where(x => !string.IsNullOrWhiteSpace(x.Descripcion) || x.ImporteBruto > 0m || x.ValorUnitario > 0m || x.Cantidad > 0m)
            .ToList();

        if (detallesValidos.Count == 0)
        {
            throw new InvalidOperationException("Debe registrar al menos un concepto.");
        }

        if (detallesValidos.Any(x => x.Cantidad <= 0m || x.ValorUnitario < 0m || x.ImporteBruto < 0m))
        {
            throw new InvalidOperationException("El detalle contiene valores no validos.");
        }

        var totalDetalle = detallesValidos.Sum(x => x.ImporteBruto);
        if (baseImponible > 0m && totalDetalle > 0m && decimal.Round(totalDetalle, 2) != decimal.Round(baseImponible, 2))
        {
            throw new InvalidOperationException("La suma del detalle debe coincidir con la base imponible.");
        }
    }

    private static decimal ResolverMonto(string componenteContable, AsientoPreviewRequest request)
    {
        return (componenteContable ?? string.Empty).Trim().ToUpperInvariant() switch
        {
            "BRUTO" => request.BaseImponible,
            "IGV" => request.Igv,
            "ISC" => request.Isc,
            "OTROS" => request.OtrosTributos,
            "REDONDEO" => request.Redondeo,
            "TOTAL" => request.ImporteTotal,
            _ => 0m
        };
    }

    private static string ConstruirReferencia(string tipoComprobante, string serie, string numero)
    {
        var tipo = string.IsNullOrWhiteSpace(tipoComprobante) ? "??" : tipoComprobante.Trim().ToUpperInvariant();
        var serieNormalizada = string.IsNullOrWhiteSpace(serie) ? "S/N" : serie.Trim().ToUpperInvariant();
        var numeroNormalizado = string.IsNullOrWhiteSpace(numero) ? "S/N" : numero.Trim().ToUpperInvariant();
        return $"{tipo} {serieNormalizada}-{numeroNormalizado}";
    }
}
