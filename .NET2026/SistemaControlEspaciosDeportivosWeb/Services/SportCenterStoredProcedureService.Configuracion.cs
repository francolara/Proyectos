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
            LogoUrl = dr.FieldCount > 13 && !dr.IsDBNull(13) ? dr.GetString(13) : null,
            PermitirModificarPrecioReserva = dr.FieldCount > 14 && !dr.IsDBNull(14) && Convert.ToBoolean(dr.GetValue(14)),
            CancelacionAutomaticaNoConfirmada = dr.FieldCount > 15 && !dr.IsDBNull(15) && Convert.ToBoolean(dr.GetValue(15)),
            MinutosCancelacionNoConfirmada = dr.FieldCount > 16 && !dr.IsDBNull(16) ? Convert.ToInt32(dr.GetValue(16)) : null,
            SedesPermitidas = dr.FieldCount > 17 && !dr.IsDBNull(17) ? Convert.ToInt32(dr.GetValue(17)) : 2,
            EspaciosPermitidos = dr.FieldCount > 18 && !dr.IsDBNull(18) ? Convert.ToInt32(dr.GetValue(18)) : 6
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
            AddParam(cmd, "@PermitirModificarPrecioReserva", model.PermitirModificarPrecioReserva, SqlDbType.Bit);
            AddParam(cmd, "@CancelacionAutomaticaNoConfirmada", model.CancelacionAutomaticaNoConfirmada, SqlDbType.Bit);
            AddParam(cmd, "@MinutosCancelacionNoConfirmada", model.MinutosCancelacionNoConfirmada, SqlDbType.Int);
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
        await using var cmdSp = new SqlCommand("Sp_ParametrosGlobales_ObtenerValor", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmdSp, "@NombreParametro", nombreParametro, SqlDbType.NVarChar);
        var resultSp = await cmdSp.ExecuteScalarAsync();
        return resultSp?.ToString();
    }

    public async Task ParametrosGlobalesUpsertValorAsync(string nombreParametro, string? descripcion, string? valorParametro, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();

        var valorNorm = (valorParametro ?? string.Empty).Trim();
        if (valorNorm.Length > 500) valorNorm = valorNorm[..500];

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

    public async Task<List<PlataformaNegocioLimiteItemViewModel>> PlataformaNegociosListarAsync(string? buscar = null)
    {
        var list = new List<PlataformaNegocioLimiteItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(
            @"SELECT n.Id, n.NombreComercial, n.Activo,
                     CAST(COALESCE(n.SedesPermitidas,2) AS INT) AS SedesPermitidas,
                     CAST(COALESCE(n.EspaciosPermitidos,6) AS INT) AS EspaciosPermitidos,
                     CAST(COALESCE(ns.EstadoSuscripcion, 0) AS INT) AS EstadoSuscripcion,
                     CAST(COALESCE(ns.EsPrueba, 0) AS BIT) AS EsPrueba,
                     ns.FechaInicioPrueba,
                     ns.FechaFinPrueba,
                     ns.TipoCobro,
                     ns.FechaInicioPlan,
                     ns.FechaFinPlan,
                     CAST(COALESCE(ns.DiasGracia, 5) AS INT) AS DiasGracia,
                     ns.FechaFinGracia
              FROM dbo.Negocios n
              LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
              WHERE (@Buscar IS NULL OR n.NombreComercial LIKE '%' + @Buscar + '%')
              ORDER BY n.NombreComercial, n.Id;", cn);
        AddParam(cmd, "@Buscar", string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(), SqlDbType.NVarChar);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new PlataformaNegocioLimiteItemViewModel
            {
                NegocioId = dr.GetInt32(0),
                NombreComercial = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Activo = !dr.IsDBNull(2) && Convert.ToBoolean(dr.GetValue(2)),
                SedesPermitidas = dr.IsDBNull(3) ? 2 : Convert.ToInt32(dr.GetValue(3)),
                EspaciosPermitidos = dr.IsDBNull(4) ? 6 : Convert.ToInt32(dr.GetValue(4)),
                EstadoSuscripcion = dr.IsDBNull(5) ? 0 : Convert.ToInt32(dr.GetValue(5)),
                EsPrueba = !dr.IsDBNull(6) && Convert.ToBoolean(dr.GetValue(6)),
                FechaInicioPrueba = dr.IsDBNull(7) ? null : dr.GetDateTime(7),
                FechaFinPrueba = dr.IsDBNull(8) ? null : dr.GetDateTime(8),
                TipoCobro = dr.IsDBNull(9) ? null : dr.GetString(9),
                FechaInicioPlan = dr.IsDBNull(10) ? null : dr.GetDateTime(10),
                FechaFinPlan = dr.IsDBNull(11) ? null : dr.GetDateTime(11),
                DiasGracia = dr.IsDBNull(12) ? 5 : Convert.ToInt32(dr.GetValue(12)),
                FechaFinGracia = dr.IsDBNull(13) ? null : dr.GetDateTime(13)
            });
            list[^1].EstadoSuscripcionNombre = list[^1].EstadoSuscripcion switch
            {
                1 when list[^1].EsPrueba => "Prueba",
                1 => "Activa",
                2 => "Contrato",
                3 => "Vencida",
                4 => "Suspendida",
                _ => "Pendiente de activacion"
            };
        }

        return list;
    }

    public async Task<bool> PlataformaNegocioActualizarLimitesAsync(int negocioId, int sedesPermitidas, int espaciosPermitidos, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(
            @"UPDATE dbo.Negocios
              SET SedesPermitidas = @SedesPermitidas,
                  EspaciosPermitidos = @EspaciosPermitidos
              WHERE Id = @NegocioId;", cn);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedesPermitidas", sedesPermitidas, SqlDbType.Int);
        AddParam(cmd, "@EspaciosPermitidos", espaciosPermitidos, SqlDbType.Int);
        var rows = await cmd.ExecuteNonQueryAsync();
        return rows > 0;
    }

    public async Task<bool> PlataformaNegocioActivarContratoAsync(int negocioId, string tipoCobro, DateOnly fechaDesde, DateOnly fechaHasta, int diasGracia, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcion_ActivarPlan", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@TipoCobro", string.IsNullOrWhiteSpace(tipoCobro) ? "MENSUAL" : tipoCobro.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@FechaInicioPlan", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@FechaFinPlan", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@DiasGracia", diasGracia <= 0 ? 5 : diasGracia, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<MiSuscripcionNegocioViewModel?> MiSuscripcionObtenerAsync(int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(
            @"SELECT TOP 1
                     ns.NegocioId,
                     CAST(COALESCE(ns.EstadoSuscripcion, 0) AS INT) AS EstadoSuscripcion,
                     CAST(COALESCE(ns.EsPrueba, 0) AS BIT) AS EsPrueba,
                     ns.FechaInicioPrueba,
                     ns.FechaFinPrueba,
                     ns.TipoCobro,
                     ns.FechaInicioPlan,
                     ns.FechaFinPlan,
                     CAST(COALESCE(ns.DiasGracia, 5) AS INT) AS DiasGracia,
                     ns.FechaFinGracia
              FROM dbo.NegociosSuscripcion ns
              WHERE ns.NegocioId = @NegocioId;", cn);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        var model = new MiSuscripcionNegocioViewModel
        {
            NegocioId = dr.GetInt32(0),
            EstadoSuscripcion = dr.IsDBNull(1) ? 0 : Convert.ToInt32(dr.GetValue(1)),
            EsPrueba = !dr.IsDBNull(2) && Convert.ToBoolean(dr.GetValue(2)),
            FechaInicioPrueba = dr.IsDBNull(3) ? null : dr.GetDateTime(3),
            FechaFinPrueba = dr.IsDBNull(4) ? null : dr.GetDateTime(4),
            TipoCobro = dr.IsDBNull(5) ? null : dr.GetString(5),
            FechaInicioPlan = dr.IsDBNull(6) ? null : dr.GetDateTime(6),
            FechaFinPlan = dr.IsDBNull(7) ? null : dr.GetDateTime(7),
            DiasGracia = dr.IsDBNull(8) ? 5 : Convert.ToInt32(dr.GetValue(8)),
            FechaFinGracia = dr.IsDBNull(9) ? null : dr.GetDateTime(9)
        };

        model.EstadoSuscripcionNombre = model.EstadoSuscripcion switch
        {
            1 when model.EsPrueba => "Prueba gratuita",
            1 => "Activa",
            2 => "Contrato activo",
            3 => "Vencida",
            4 => "Suspendida",
            _ => "Pendiente de activacion"
        };

        return model;
    }
}
