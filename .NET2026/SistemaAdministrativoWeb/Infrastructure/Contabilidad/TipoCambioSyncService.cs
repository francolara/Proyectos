namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoCambioSyncService(
    ITipoCambioRepository tipoCambioRepository,
    IMigoTipoCambioApiClient migoTipoCambioApiClient) : ITipoCambioSyncService
{
    public async Task<TipoCambioDto?> SincronizarFechaAsync(int idCuentaAdministradora, DateOnly fecha, string idMoneda, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        var codigoMoneda = NormalizarMoneda(idMoneda);
        if (!string.Equals(codigoMoneda, "USD", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("La integracion de Migo para tipos de cambio solo devuelve cotizacion USD.");
        }

        var apiItem = await migoTipoCambioApiClient.ObtenerPorFechaAsync(fecha, cancellationToken);
        if (apiItem is null || !string.Equals(apiItem.Moneda, codigoMoneda, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var existente = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(idCuentaAdministradora, fecha, codigoMoneda, cancellationToken);
        return await tipoCambioRepository.GuardarAsync(new GuardarTipoCambioRequest
        {
            IdTipoCambio = existente?.IdTipoCambio,
            IdCuentaAdministradora = idCuentaAdministradora,
            Fecha = apiItem.Fecha,
            IdMoneda = codigoMoneda,
            Compra = apiItem.PrecioCompra,
            Venta = apiItem.PrecioVenta,
            CompraSbs = apiItem.PrecioCompra,
            VentaSbs = apiItem.PrecioVenta,
            Fuente = "API",
            Estado = true,
            UsuarioRegistro = usuarioRegistro
        }, cancellationToken);
    }

    public async Task<TipoCambioPeriodoSyncResult> SincronizarPeriodoAsync(int idCuentaAdministradora, short anio, byte mes, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        var fechaInicio = new DateOnly(anio, mes, 1);
        var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);
        var apiItems = await migoTipoCambioApiClient.ObtenerPorRangoAsync(fechaInicio, fechaFin, cancellationToken);
        var sincronizados = new List<TipoCambioDto>();

        foreach (var apiItem in apiItems)
        {
            if (!string.Equals(apiItem.Moneda, "USD", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var existente = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(idCuentaAdministradora, apiItem.Fecha, apiItem.Moneda, cancellationToken);
            var guardado = await tipoCambioRepository.GuardarAsync(new GuardarTipoCambioRequest
            {
                IdTipoCambio = existente?.IdTipoCambio,
                IdCuentaAdministradora = idCuentaAdministradora,
                Fecha = apiItem.Fecha,
                IdMoneda = apiItem.Moneda,
                Compra = apiItem.PrecioCompra,
                Venta = apiItem.PrecioVenta,
                CompraSbs = apiItem.PrecioCompra,
                VentaSbs = apiItem.PrecioVenta,
                Fuente = "API",
                Estado = true,
                UsuarioRegistro = usuarioRegistro
            }, cancellationToken);

            sincronizados.Add(guardado);
        }

        return new TipoCambioPeriodoSyncResult
        {
            TotalConsultados = apiItems.Count,
            TotalSincronizados = sincronizados.Count,
            TiposCambio = sincronizados
        };
    }

    private static string NormalizarMoneda(string idMoneda)
    {
        return (idMoneda ?? string.Empty).Trim().ToUpperInvariant();
    }
}
