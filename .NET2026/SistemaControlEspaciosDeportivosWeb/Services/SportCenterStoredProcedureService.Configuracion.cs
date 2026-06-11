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
            EnviarComprobanteAutomatico = dr.FieldCount > 11 && !dr.IsDBNull(11) && Convert.ToBoolean(dr.GetValue(11)),
            EmisionReciboInterno = dr.FieldCount > 12 && !dr.IsDBNull(12) && Convert.ToBoolean(dr.GetValue(12)),
            PorcentajeIgv = dr.FieldCount > 13 && !dr.IsDBNull(13) ? Convert.ToInt32(dr.GetValue(13)) : 18,
            LogoUrl = dr.FieldCount > 14 && !dr.IsDBNull(14) ? dr.GetString(14) : null,
            PermitirModificarPrecioReserva = dr.FieldCount > 15 && !dr.IsDBNull(15) && Convert.ToBoolean(dr.GetValue(15)),
            CancelacionAutomaticaNoConfirmada = dr.FieldCount > 16 && !dr.IsDBNull(16) && Convert.ToBoolean(dr.GetValue(16)),
            MinutosCancelacionNoConfirmada = dr.FieldCount > 17 && !dr.IsDBNull(17) ? Convert.ToInt32(dr.GetValue(17)) : null,
            SedesPermitidas = dr.FieldCount > 18 && !dr.IsDBNull(18) ? Convert.ToInt32(dr.GetValue(18)) : 2,
            EspaciosPermitidos = dr.FieldCount > 19 && !dr.IsDBNull(19) ? Convert.ToInt32(dr.GetValue(19)) : 6,
            HorasMaximasReservaCliente = dr.FieldCount > 20 && !dr.IsDBNull(20) ? Convert.ToInt32(dr.GetValue(20)) : 1
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
            AddParam(cmd, "@EnviarComprobanteAutomatico", model.EnviarComprobanteAutomatico, SqlDbType.Bit);
            AddParam(cmd, "@EmisionReciboInterno", model.EmisionReciboInterno, SqlDbType.Bit);
            AddParam(cmd, "@PorcentajeIgv", model.PorcentajeIgv, SqlDbType.Int);
            AddParam(cmd, "@LogoUrl", model.LogoUrl, SqlDbType.NVarChar);
            AddParam(cmd, "@PermitirModificarPrecioReserva", model.PermitirModificarPrecioReserva, SqlDbType.Bit);
            AddParam(cmd, "@CancelacionAutomaticaNoConfirmada", model.CancelacionAutomaticaNoConfirmada, SqlDbType.Bit);
            AddParam(cmd, "@MinutosCancelacionNoConfirmada", model.MinutosCancelacionNoConfirmada, SqlDbType.Int);
            AddParam(cmd, "@HorasMaximasReservaCliente", model.HorasMaximasReservaCliente, SqlDbType.Int);
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

    public async Task<bool> ConfiguracionClubActualizarEmisionAsync(int negocioId, bool emisionComprobantesElectronicos, bool enviarComprobanteAutomatico, bool emisionReciboInterno, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_ConfiguracionClub_ActualizarEmision", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@EmisionComprobantesElectronicos", emisionComprobantesElectronicos, SqlDbType.Bit);
            AddParam(cmd, "@EnviarComprobanteAutomatico", enviarComprobanteAutomatico, SqlDbType.Bit);
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

    public async Task<(List<PlataformaNegocioLimiteItemViewModel> Negocios, int TotalRegistros)> PlataformaNegociosListarAsync(string? buscar = null, string? estadoContrato = null, int pagina = 1, int tamanoPagina = 20)
    {
        var list = new List<PlataformaNegocioLimiteItemViewModel>();
        var paginaNormalizada = pagina < 1 ? 1 : pagina;
        var tamanoNormalizado = tamanoPagina < 1 ? 20 : tamanoPagina;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Plataforma_Negocios_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Buscar", string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(), SqlDbType.NVarChar);
        AddParam(cmd, "@EstadoContrato", string.IsNullOrWhiteSpace(estadoContrato) ? "todos" : estadoContrato.Trim().ToLowerInvariant(), SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", paginaNormalizada, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoNormalizado, SqlDbType.Int);
        var totalRegistrosParam = cmd.Parameters.Add("@TotalRegistros", SqlDbType.Int);
        totalRegistrosParam.Direction = ParameterDirection.Output;
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
                UsuariosPermitidos = dr.IsDBNull(5) ? 3 : Convert.ToInt32(dr.GetValue(5)),
                EstadoSuscripcion = dr.IsDBNull(6) ? 0 : Convert.ToInt32(dr.GetValue(6)),
                EsPrueba = !dr.IsDBNull(7) && Convert.ToBoolean(dr.GetValue(7)),
                FechaInicioPrueba = dr.IsDBNull(8) ? null : dr.GetDateTime(8),
                FechaFinPrueba = dr.IsDBNull(9) ? null : dr.GetDateTime(9),
                TipoCobro = dr.IsDBNull(10) ? null : dr.GetString(10),
                FechaInicioPlan = dr.IsDBNull(11) ? null : dr.GetDateTime(11),
                FechaFinPlan = dr.IsDBNull(12) ? null : dr.GetDateTime(12),
                DiasGracia = dr.IsDBNull(13) ? 5 : Convert.ToInt32(dr.GetValue(13)),
                FechaFinGracia = dr.IsDBNull(14) ? null : dr.GetDateTime(14)
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
        await dr.CloseAsync();

        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        return (list, totalRegistros);
    }

public async Task<(string? Correo, string? NombreDestino, string? Telefono)> PlataformaNegocioObtenerContactoCorreoAsync(int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();

        const string sql = @"
SELECT TOP (1)
    u.Email AS Correo,
    COALESCE(
        NULLIF(LTRIM(RTRIM(u.PhoneNumber)), N''),
        NULLIF(LTRIM(RTRIM(sref.Telefono)), N''),
        NULLIF(LTRIM(RTRIM(acref.Telefono)), N'')
    ) AS Telefono,
    COALESCE(
        NULLIF(LTRIM(RTRIM(CONCAT(COALESCE(u.Nombres, N''), N' ', COALESCE(u.Apellidos, N'')))), N''),
        NULLIF(LTRIM(RTRIM(u.UserName)), N''),
        NULLIF(LTRIM(RTRIM(n.NombreComercial)), N''),
        N'Cliente') AS NombreDestino
FROM dbo.UsuariosNegocio un
INNER JOIN dbo.AspNetUsers u ON u.Id = un.UsuarioId
INNER JOIN dbo.Negocios n ON n.Id = un.NegocioId
OUTER APPLY (
    SELECT TOP (1) s.Telefono
    FROM dbo.Sedes s
    WHERE s.NegocioId = n.Id
      AND s.Telefono IS NOT NULL
      AND LTRIM(RTRIM(s.Telefono)) <> N''
    ORDER BY s.Id ASC
) sref
OUTER APPLY (
    SELECT TOP (1) ac.Telefono
    FROM dbo.SolicitudesAltaClub ac
    WHERE ac.NegocioId = n.Id
      AND ac.Telefono IS NOT NULL
      AND LTRIM(RTRIM(ac.Telefono)) <> N''
    ORDER BY ac.Id DESC
) acref
WHERE un.NegocioId = @NegocioId
  AND u.Email IS NOT NULL
  AND LTRIM(RTRIM(u.Email)) <> N''
ORDER BY un.Id ASC;

IF @@ROWCOUNT = 0
BEGIN
SELECT TOP (1)
        scn.CorreoNotificacion AS Correo,
        COALESCE(
            NULLIF(LTRIM(RTRIM(s.Telefono)), N''),
            NULLIF(LTRIM(RTRIM(acref.Telefono)), N'')
        ) AS Telefono,
        COALESCE(NULLIF(LTRIM(RTRIM(n.NombreComercial)), N''), N'Cliente') AS NombreDestino
    FROM dbo.Sedes s
    INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
    INNER JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
    OUTER APPLY (
        SELECT TOP (1) ac.Telefono
        FROM dbo.SolicitudesAltaClub ac
        WHERE ac.NegocioId = n.Id
          AND ac.Telefono IS NOT NULL
          AND LTRIM(RTRIM(ac.Telefono)) <> N''
        ORDER BY ac.Id DESC
    ) acref
    WHERE s.NegocioId = @NegocioId
      AND scn.CorreoNotificacion IS NOT NULL
      AND LTRIM(RTRIM(scn.CorreoNotificacion)) <> N''
    ORDER BY s.Id ASC;
END";

        await using var cmd = new SqlCommand(sql, cn);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
    if (!await dr.ReadAsync())
        return (null, null, null);

    var correo = dr.IsDBNull(0) ? null : dr.GetString(0)?.Trim();
    var telefono = dr.IsDBNull(1) ? null : dr.GetString(1)?.Trim();
    var nombre = dr.IsDBNull(2) ? null : dr.GetString(2)?.Trim();
    return (
        string.IsNullOrWhiteSpace(correo) ? null : correo,
        string.IsNullOrWhiteSpace(nombre) ? null : nombre,
        string.IsNullOrWhiteSpace(telefono) ? null : telefono);
}

    public async Task<bool> PlataformaNegocioActualizarLimitesAsync(int negocioId, int sedesPermitidas, int espaciosPermitidos, int usuariosPermitidos, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Plataforma_Negocios_ActualizarLimites", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@SedesPermitidas", sedesPermitidas, SqlDbType.Int);
            AddParam(cmd, "@EspaciosPermitidos", espaciosPermitidos, SqlDbType.Int);
            AddParam(cmd, "@UsuariosPermitidos", usuariosPermitidos, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<(int SedesPermitidas, int EspaciosPermitidos, int UsuariosPermitidos)> NegocioObtenerLimitesOperativosAsync(int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(
            @"SELECT
                  CAST(COALESCE(SedesPermitidas, 2) AS INT) AS SedesPermitidas,
                  CAST(COALESCE(EspaciosPermitidos, 6) AS INT) AS EspaciosPermitidos,
                  CAST(COALESCE(UsuariosPermitidos, 3) AS INT) AS UsuariosPermitidos
              FROM dbo.Negocios
              WHERE Id = @NegocioId;", cn);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
            return (2, 6, 3);

        return (
            dr.IsDBNull(0) ? 2 : Convert.ToInt32(dr.GetValue(0)),
            dr.IsDBNull(1) ? 6 : Convert.ToInt32(dr.GetValue(1)),
            dr.IsDBNull(2) ? 3 : Convert.ToInt32(dr.GetValue(2))
        );
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

    public async Task<bool> PlataformaNegocioRenovarContratoAsync(int negocioId, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcion_RenovarPlan", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> PlataformaNegocioFinalizarContratoAsync(int negocioId, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcion_FinalizarPlan", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> PlataformaNegocioExtenderPruebaAsync(int negocioId, int diasExtra, string? observacion, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcion_ExtenderPrueba", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@DiasExtra", diasExtra, SqlDbType.Int);
            AddParam(cmd, "@Observacion", string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> PlataformaNegocioAplicarGraciaManualAsync(int negocioId, int diasExtra, string? observacion, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcion_AplicarGraciaManual", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@DiasExtra", diasExtra, SqlDbType.Int);
            AddParam(cmd, "@Observacion", string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> PlataformaNegocioCambiarPlanAsync(int negocioId, string tipoCobro, DateOnly fechaDesde, DateOnly fechaHasta, int diasGracia, string? observacion, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcion_CambiarPlan", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@TipoCobro", string.IsNullOrWhiteSpace(tipoCobro) ? "MENSUAL" : tipoCobro.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@FechaInicioPlan", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@FechaFinPlan", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@DiasGracia", diasGracia <= 0 ? 5 : diasGracia, SqlDbType.Int);
            AddParam(cmd, "@Observacion", string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<List<PlataformaNegocioSuscripcionMovimientoViewModel>> PlataformaNegocioHistorialComercialAsync(int negocioId, int top = 8)
    {
        var list = new List<PlataformaNegocioSuscripcionMovimientoViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_NegociosSuscripcionMovimiento_ListarPorNegocio", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Top", top <= 0 ? 8 : top, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            var item = new PlataformaNegocioSuscripcionMovimientoViewModel
            {
                Id = dr.IsDBNull(0) ? 0 : Convert.ToInt32(dr.GetValue(0)),
                TipoMovimiento = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                EstadoSuscripcionAnterior = dr.IsDBNull(2) ? 0 : Convert.ToInt32(dr.GetValue(2)),
                EstadoSuscripcionNuevo = dr.IsDBNull(3) ? 0 : Convert.ToInt32(dr.GetValue(3)),
                EsPruebaAnterior = !dr.IsDBNull(4) && Convert.ToBoolean(dr.GetValue(4)),
                EsPruebaNuevo = !dr.IsDBNull(5) && Convert.ToBoolean(dr.GetValue(5)),
                TipoCobroAnterior = dr.IsDBNull(6) ? null : dr.GetString(6),
                TipoCobroNuevo = dr.IsDBNull(7) ? null : dr.GetString(7),
                FechaInicioReferencia = dr.IsDBNull(8) ? null : dr.GetDateTime(8),
                FechaFinReferencia = dr.IsDBNull(9) ? null : dr.GetDateTime(9),
                DiasGracia = dr.IsDBNull(10) ? 0 : Convert.ToInt32(dr.GetValue(10)),
                DiasExtra = dr.IsDBNull(11) ? 0 : Convert.ToInt32(dr.GetValue(11)),
                Observacion = dr.IsDBNull(12) ? null : dr.GetString(12),
                FechaCreacion = dr.IsDBNull(13) ? DateTime.MinValue : dr.GetDateTime(13),
                UsuarioCreacion = dr.IsDBNull(14) ? null : dr.GetString(14)
            };

            item.TipoMovimientoNombre = item.TipoMovimiento switch
            {
                "ALTA_PRUEBA" => "Alta de prueba",
                "EXTENSION_PRUEBA" => "Extension de prueba",
                "ACTIVACION_CONTRATO" => "Activacion de contrato",
                "RENOVACION" => "Renovacion",
                "CAMBIO_PLAN" => "Cambio de plan",
                "GRACIA_MANUAL" => "Gracia manual",
                "FINALIZACION" => "Finalizacion",
                _ => item.TipoMovimiento
            };

            list.Add(item);
        }

        return list;
    }

    public async Task<bool> PlataformaNegocioRegistrarPagoSuscripcionAsync(int negocioId, string tipoPago, string estadoPago, decimal monto, string moneda, DateTime fechaPago, DateOnly? fechaVencimiento, string? operacionNumero, string? entidadFinanciera, string? referenciaExterna, string? observacion, string? accionAplicacion, bool aplicarAlConfirmar, string? tipoCobroObjetivo, DateOnly? fechaInicioPlanObjetivo, int? diasGraciaObjetivo, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcionPago_Registrar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@TipoPago", string.IsNullOrWhiteSpace(tipoPago) ? "TRANSFERENCIA" : tipoPago.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@EstadoPago", string.IsNullOrWhiteSpace(estadoPago) ? "PAGADO" : estadoPago.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@Monto", monto, SqlDbType.Decimal);
            cmd.Parameters["@Monto"].Precision = 12;
            cmd.Parameters["@Monto"].Scale = 2;
            AddParam(cmd, "@Moneda", string.IsNullOrWhiteSpace(moneda) ? "PEN" : moneda.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@FechaPago", fechaPago, SqlDbType.DateTime2);
            AddParam(cmd, "@FechaVencimiento", fechaVencimiento?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@OperacionNumero", string.IsNullOrWhiteSpace(operacionNumero) ? null : operacionNumero.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@EntidadFinanciera", string.IsNullOrWhiteSpace(entidadFinanciera) ? null : entidadFinanciera.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@ReferenciaExterna", string.IsNullOrWhiteSpace(referenciaExterna) ? null : referenciaExterna.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@Observacion", string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@AccionAplicacion", string.IsNullOrWhiteSpace(accionAplicacion) ? null : accionAplicacion.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@AplicarAlConfirmar", aplicarAlConfirmar, SqlDbType.Bit);
            AddParam(cmd, "@TipoCobroObjetivo", string.IsNullOrWhiteSpace(tipoCobroObjetivo) ? null : tipoCobroObjetivo.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@FechaInicioPlanObjetivo", fechaInicioPlanObjetivo?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@DiasGraciaObjetivo", diasGraciaObjetivo, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<(List<PlataformaNegocioSuscripcionPagoViewModel> Pagos, int CantidadPagos, decimal MontoTotalPagado, DateTime? UltimaFechaPago, decimal? UltimoMonto, string? UltimoTipoPago)> PlataformaNegocioPagosSuscripcionAsync(int negocioId, int top = 8)
    {
        var list = new List<PlataformaNegocioSuscripcionPagoViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_NegociosSuscripcionPago_ListarPorNegocio", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Top", top <= 0 ? 8 : top, SqlDbType.Int);

        var cantidadParam = cmd.Parameters.Add("@CantidadPagos", SqlDbType.Int);
        cantidadParam.Direction = ParameterDirection.Output;

        var totalParam = cmd.Parameters.Add("@MontoTotalPagado", SqlDbType.Decimal);
        totalParam.Direction = ParameterDirection.Output;
        totalParam.Precision = 12;
        totalParam.Scale = 2;

        var ultimaFechaParam = cmd.Parameters.Add("@UltimaFechaPago", SqlDbType.DateTime2);
        ultimaFechaParam.Direction = ParameterDirection.Output;

        var ultimoMontoParam = cmd.Parameters.Add("@UltimoMonto", SqlDbType.Decimal);
        ultimoMontoParam.Direction = ParameterDirection.Output;
        ultimoMontoParam.Precision = 12;
        ultimoMontoParam.Scale = 2;

        var ultimoTipoParam = cmd.Parameters.Add("@UltimoTipoPago", SqlDbType.NVarChar, 30);
        ultimoTipoParam.Direction = ParameterDirection.Output;

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new PlataformaNegocioSuscripcionPagoViewModel
            {
                Id = dr.IsDBNull(0) ? 0 : Convert.ToInt32(dr.GetValue(0)),
                TipoPago = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                EstadoPago = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
                Monto = dr.IsDBNull(3) ? 0 : dr.GetDecimal(3),
                Moneda = dr.IsDBNull(4) ? "PEN" : dr.GetString(4),
                FechaPago = dr.IsDBNull(5) ? DateTime.MinValue : dr.GetDateTime(5),
                FechaVencimiento = dr.IsDBNull(6) ? null : dr.GetDateTime(6),
                OperacionNumero = dr.IsDBNull(7) ? null : dr.GetString(7),
                EntidadFinanciera = dr.IsDBNull(8) ? null : dr.GetString(8),
                ReferenciaExterna = dr.IsDBNull(9) ? null : dr.GetString(9),
                Observacion = dr.IsDBNull(10) ? null : dr.GetString(10),
                FechaCreacion = dr.IsDBNull(11) ? DateTime.MinValue : dr.GetDateTime(11),
                UsuarioCreacion = dr.IsDBNull(12) ? null : dr.GetString(12),
                TipoMovimientoRelacionado = dr.IsDBNull(13) ? null : dr.GetString(13),
                AccionAplicacion = dr.IsDBNull(14) ? null : dr.GetString(14),
                AplicarAlConfirmar = !dr.IsDBNull(15) && Convert.ToBoolean(dr.GetValue(15)),
                AplicadoSuscripcion = !dr.IsDBNull(16) && Convert.ToBoolean(dr.GetValue(16)),
                FechaAplicacion = dr.IsDBNull(17) ? null : dr.GetDateTime(17),
                UsuarioAplicacion = dr.IsDBNull(18) ? null : dr.GetString(18),
                TipoCobroObjetivo = dr.IsDBNull(19) ? null : dr.GetString(19),
                FechaInicioPlanObjetivo = dr.IsDBNull(20) ? null : dr.GetDateTime(20),
                DiasGraciaObjetivo = dr.IsDBNull(21) ? null : Convert.ToInt32(dr.GetValue(21))
            });
        }
        await dr.CloseAsync();

        return (
            list,
            cantidadParam.Value is int cantidad ? cantidad : 0,
            totalParam.Value == DBNull.Value ? 0 : Convert.ToDecimal(totalParam.Value),
            ultimaFechaParam.Value == DBNull.Value ? null : Convert.ToDateTime(ultimaFechaParam.Value),
            ultimoMontoParam.Value == DBNull.Value ? null : Convert.ToDecimal(ultimoMontoParam.Value),
            ultimoTipoParam.Value == DBNull.Value ? null : Convert.ToString(ultimoTipoParam.Value)
        );
    }

    public async Task<bool> PlataformaNegocioConfirmarPagoSuscripcionAsync(int negocioId, int pagoId, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_NegociosSuscripcionPago_ConfirmarAplicar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@PagoId", pagoId, SqlDbType.Int);
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
