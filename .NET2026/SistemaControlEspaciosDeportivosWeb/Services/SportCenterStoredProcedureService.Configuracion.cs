using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<ConfiguracionClubViewModel?> ConfiguracionClubObtenerAsync(int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_ConfiguracionClub_Obtener", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new ConfiguracionClubViewModel
        {
            Id = dr.GetInt32(0),
            NegocioId = dr.GetInt32(0),
            NombreComercial = dr.GetString(1),
            RazonSocial = dr.IsDBNull(2) ? null : dr.GetString(2),
            TipoDocumento = dr.IsDBNull(3) ? "1" : dr.GetString(3),
            NumeroDocumento = dr.IsDBNull(4) ? null : dr.GetString(4),
            DireccionFiscal = dr.IsDBNull(5) ? null : dr.GetString(5),
            MonedaId = dr.IsDBNull(6) ? 1 : dr.GetInt32(6),
            CodigoUbigeo = dr.IsDBNull(7) ? null : dr.GetString(7),
            PoliticaConfirmacionPago = dr.FieldCount > 8 && !dr.IsDBNull(8) ? Convert.ToInt32(dr.GetValue(8)) : 0,
            PorcentajeAdelantoMinimo = dr.FieldCount > 9 && !dr.IsDBNull(9) ? Convert.ToDecimal(dr.GetValue(9)) : null,
            EmisionComprobantesElectronicos = dr.FieldCount > 10 && !dr.IsDBNull(10) && Convert.ToBoolean(dr.GetValue(10)),
            EmisionReciboInterno = dr.FieldCount > 11 && !dr.IsDBNull(11) && Convert.ToBoolean(dr.GetValue(11)),
            PorcentajeIgv = dr.FieldCount > 12 && !dr.IsDBNull(12) ? Convert.ToInt32(dr.GetValue(12)) : 18,
            LogoUrl = dr.FieldCount > 13 && !dr.IsDBNull(13) ? dr.GetString(13) : null
        };
    }

    public async Task<bool> ConfiguracionClubActualizarAsync(ConfiguracionClubViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_ConfiguracionClub_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@NombreComercial", model.NombreComercial, SqlDbType.NVarChar);
            AddParam(cmd, "@RazonSocial", model.RazonSocial, SqlDbType.NVarChar);
            AddParam(cmd, "@TipoDocumentoFiscal", model.TipoDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@NumeroDocumentoFiscal", model.NumeroDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@DireccionFiscal", model.DireccionFiscal, SqlDbType.NVarChar);
            AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
            AddParam(cmd, "@MonedaId", model.MonedaId, SqlDbType.Int);
            AddParam(cmd, "@PoliticaConfirmacionPago", model.PoliticaConfirmacionPago, SqlDbType.TinyInt);
            AddParam(cmd, "@PorcentajeAdelantoMinimo", model.PorcentajeAdelantoMinimo, SqlDbType.Decimal);
            AddParam(cmd, "@EmisionComprobantesElectronicos", model.EmisionComprobantesElectronicos, SqlDbType.Bit);
            AddParam(cmd, "@EmisionReciboInterno", model.EmisionReciboInterno, SqlDbType.Bit);
            AddParam(cmd, "@PorcentajeIgv", model.PorcentajeIgv, SqlDbType.Int);
            AddParam(cmd, "@LogoUrl", model.LogoUrl, SqlDbType.NVarChar);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<List<SelectListItem>> ConfiguracionClubComboMonedasAsync(int negocioId)
    {
        var list = new List<SelectListItem>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Combos_Monedas", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SelectListItem(dr.GetString(1), dr.GetInt32(0).ToString()));
        }
        return list;
    }

    public async Task<bool> ConfiguracionClubActualizarEmisionAsync(int negocioId, bool emisionComprobantesElectronicos, bool emisionReciboInterno, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_ConfiguracionClub_ActualizarEmision", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@EmisionComprobantesElectronicos", emisionComprobantesElectronicos, SqlDbType.Bit);
            AddParam(cmd, "@EmisionReciboInterno", emisionReciboInterno, SqlDbType.Bit);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<List<SelectListItem>> CombosTiposDocumentoIdentidadSunatAsync()
    {
        var list = new List<SelectListItem>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Combos_TiposDocumentoIdentidadSunat", cn) { CommandType = CommandType.StoredProcedure };
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SelectListItem(dr.GetString(1), dr.GetString(0)));
        }

        return list;
    }

    public Task<List<SelectListItem>> CombosTiposNotaComprobanteSunatAsync(string tipoNota)
        => ComboAsync(
            "Sp_Combos_TiposNotaComprobanteSunat",
            ("@TipoNota", string.IsNullOrWhiteSpace(tipoNota) ? null : tipoNota.Trim().ToUpperInvariant(), SqlDbType.Char));

    public Task<List<SelectListItem>> CombosDocumentosComprobanteNegocioAsync(int negocioId, bool? tributario = null)
        => ComboAsync(
            "Sp_Combos_DocumentosComprobanteNegocio",
            ("@NegocioId", negocioId, SqlDbType.Int),
            ("@Tributario", tributario, SqlDbType.Bit));

    public Task<List<SelectListItem>> CombosSeriesDocumentoComprobanteAsync(int negocioId, string codigoSunat)
        => ComboAsync(
            "Sp_Combos_SeriesDocumentoComprobante",
            ("@NegocioId", negocioId, SqlDbType.Int),
            ("@CodigoSunat", codigoSunat, SqlDbType.NVarChar));

    public async Task<string?> ParametrosGlobalesObtenerValorAsync(string nombreParametro)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        try
        {
            await using var cmdSp = new SqlCommand("Sp_ParametrosGlobales_ObtenerValor", cn)
            {
                CommandType = CommandType.StoredProcedure
            };
            AddParam(cmdSp, "@NombreParametro", nombreParametro, SqlDbType.NVarChar);
            var resultSp = await cmdSp.ExecuteScalarAsync();
            return resultSp?.ToString();
        }
        catch (SqlException ex) when (ex.Message.Contains("Sp_ParametrosGlobales_ObtenerValor", StringComparison.OrdinalIgnoreCase))
        {
            await using var cmdFallback = new SqlCommand(@"
                SELECT TOP (1) p.ValorParametro
                FROM dbo.ParametrosGlobales p
                WHERE p.NombreParametro = @NombreParametro;", cn);
            AddParam(cmdFallback, "@NombreParametro", nombreParametro, SqlDbType.NVarChar);
            var resultFallback = await cmdFallback.ExecuteScalarAsync();
            return resultFallback?.ToString();
        }
    }

    public async Task ParametrosGlobalesUpsertValorAsync(string nombreParametro, string? descripcion, string? valorParametro, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();

        var valorNorm = (valorParametro ?? string.Empty).Trim();
        if (valorNorm.Length > 100) valorNorm = valorNorm[..100];

        await using var cmdSp = new SqlCommand("Sp_ParametrosGlobales_UpsertValor", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmdSp, "@NombreParametro", nombreParametro, SqlDbType.NVarChar);
        AddParam(cmdSp, "@Descripcion", descripcion, SqlDbType.NVarChar);
        AddParam(cmdSp, "@ValorParametro", valorNorm, SqlDbType.NVarChar);
        AddParam(cmdSp, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmdSp.ExecuteNonQueryAsync();
    }

    public async Task<List<SerieDocumentoComprobanteItemViewModel>> ConfiguracionSeriesDocumentoListarAsync(int negocioId)
    {
        var list = new List<SerieDocumentoComprobanteItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Configuracion_SeriesDocumentoComprobante_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SerieDocumentoComprobanteItemViewModel
            {
                Id = dr.GetInt32(0),
                CodigoSunat = dr.GetString(1),
                NombreDocumento = dr.GetString(2),
                Tributario = dr.GetBoolean(3),
                Serie = dr.GetString(4),
                Activo = dr.GetBoolean(5)
            });
        }

        return list;
    }

    public async Task ConfiguracionSeriesDocumentoGuardarAsync(int negocioId, string codigoSunat, string serie, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Configuracion_SeriesDocumentoComprobante_Guardar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@CodigoSunat", codigoSunat, SqlDbType.NVarChar);
        AddParam(cmd, "@Serie", serie, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
    }

    public async Task<bool> ConfiguracionSeriesDocumentoEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Configuracion_SeriesDocumentoComprobante_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }
}
