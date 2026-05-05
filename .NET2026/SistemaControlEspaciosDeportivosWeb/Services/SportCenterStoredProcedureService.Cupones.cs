using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<CuponValidacionViewModel> CuponesValidarAsync(int negocioId, int? sedeId, int espacioDeportivoId, string? codigoCupon, decimal montoBase)
    {
        var result = new CuponValidacionViewModel
        {
            EsValido = false,
            Mensaje = "Ingresa un cupon para validar.",
            CodigoCupon = string.Empty,
            MontoAntes = montoBase < 0 ? 0 : montoBase,
            MontoDescuento = 0,
            MontoFinal = montoBase < 0 ? 0 : montoBase
        };

        var codigoNormalizado = string.IsNullOrWhiteSpace(codigoCupon) ? null : codigoCupon.Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(codigoNormalizado))
            return result;

        if (result.MontoAntes <= 0)
        {
            result.Mensaje = "El monto base de la reserva debe ser mayor a cero.";
            return result;
        }

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(@"
            DECLARE @SedeIdCalc INT = @SedeId;
            IF @SedeIdCalc IS NULL
            BEGIN
                SELECT TOP (1) @SedeIdCalc = e.SedeId
                FROM dbo.EspaciosDeportivos e
                WHERE e.Id = @EspacioDeportivoId;
            END
            DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);
            SELECT TOP (1)
                c.CodigoCupon,
                c.TipoDescuento,
                c.ValorDescuento
            FROM dbo.Cupones c
            WHERE c.NegocioId = @NegocioId
              AND c.CodigoCupon = @CodigoCupon
              AND c.Activo = 1
              AND c.FechaInicio <= @Hoy
              AND c.FechaFin >= @Hoy
              AND c.CantidadUsosActuales < c.CantidadMaxUsos
              AND (c.SedeId IS NULL OR c.SedeId = @SedeIdCalc)
              AND (c.EspacioDeportivoId IS NULL OR c.EspacioDeportivoId = @EspacioDeportivoId);", cn);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@CodigoCupon", codigoNormalizado, SqlDbType.NVarChar);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
        {
            result.Mensaje = "Cupon invalido, vencido o sin usos disponibles para esta sede/espacio.";
            return result;
        }

        var tipo = dr.GetString(1);
        var valor = dr.GetDecimal(2);
        var descuento = tipo == "PORCENTAJE"
            ? Math.Round(result.MontoAntes * (valor / 100m), 2)
            : valor;
        if (descuento < 0) descuento = 0;
        if (descuento > result.MontoAntes) descuento = result.MontoAntes;

        result.EsValido = true;
        result.CodigoCupon = dr.GetString(0);
        result.MontoDescuento = descuento;
        result.MontoFinal = result.MontoAntes - descuento;
        result.Mensaje = $"Cupon aplicado correctamente: -{descuento:N2}.";
        return result;
    }

    public async Task<(List<CuponItemViewModel> Cupones, int TotalRegistros)> CuponesListarAsync(
        int negocioId,
        int? sedeId = null,
        DateOnly? fechaDesde = null,
        DateOnly? fechaHasta = null,
        string? estado = "vigentes",
        int pagina = 1,
        int tamanoPagina = 20)
    {
        var list = new List<CuponItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Cupones_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@Estado", string.IsNullOrWhiteSpace(estado) ? "vigentes" : estado.Trim().ToLowerInvariant(), SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", pagina < 1 ? 1 : pagina, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoPagina < 1 ? 20 : tamanoPagina, SqlDbType.Int);
        var totalRegistrosParam = new SqlParameter("@TotalRegistros", SqlDbType.Int) { Direction = ParameterDirection.Output };
        cmd.Parameters.Add(totalRegistrosParam);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new CuponItemViewModel
            {
                Id = dr.GetInt32(0),
                CodigoCupon = dr.GetString(1),
                Nombre = dr.GetString(2),
                TipoDescuento = dr.GetString(3),
                ValorDescuento = dr.GetDecimal(4),
                CantidadMaxUsos = dr.GetInt32(5),
                CantidadUsosActuales = dr.GetInt32(6),
                CantidadUsosDisponibles = dr.GetInt32(7),
                FechaInicio = DateOnly.FromDateTime(dr.GetDateTime(8)),
                FechaFin = DateOnly.FromDateTime(dr.GetDateTime(9)),
                Sede = dr.GetString(10),
                Espacio = dr.GetString(11),
                Activo = dr.GetBoolean(12),
                VigenteHoy = dr.GetBoolean(13)
            });
        }

        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        return (list, totalRegistros);
    }

    public async Task<CuponFormViewModel?> CuponesObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Cupones_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new CuponFormViewModel
        {
            Id = dr.GetInt32(0),
            SedeId = dr.IsDBNull(1) ? null : dr.GetInt32(1),
            EspacioDeportivoId = dr.IsDBNull(2) ? null : dr.GetInt32(2),
            CodigoCupon = dr.GetString(3),
            Nombre = dr.GetString(4),
            TipoDescuento = dr.GetString(5),
            ValorDescuento = dr.GetDecimal(6),
            CantidadMaxUsos = dr.GetInt32(7),
            FechaInicio = DateOnly.FromDateTime(dr.GetDateTime(8)),
            FechaFin = DateOnly.FromDateTime(dr.GetDateTime(9)),
            Activo = dr.GetBoolean(10),
            NegocioId = negocioId
        };
    }

    public async Task<int> CuponesCrearAsync(CuponFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Cupones_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", model.SedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@CodigoCupon", model.CodigoCupon.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoDescuento", model.TipoDescuento, SqlDbType.NVarChar);
        AddParam(cmd, "@ValorDescuento", model.ValorDescuento, SqlDbType.Decimal);
        AddParam(cmd, "@CantidadMaxUsos", model.CantidadMaxUsos, SqlDbType.Int);
        AddParam(cmd, "@FechaInicio", model.FechaInicio.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaFin", model.FechaFin.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> CuponesActualizarAsync(CuponFormViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Cupones_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@SedeId", model.SedeId, SqlDbType.Int);
            AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
            AddParam(cmd, "@CodigoCupon", model.CodigoCupon.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
            AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
            AddParam(cmd, "@TipoDescuento", model.TipoDescuento, SqlDbType.NVarChar);
            AddParam(cmd, "@ValorDescuento", model.ValorDescuento, SqlDbType.Decimal);
            AddParam(cmd, "@CantidadMaxUsos", model.CantidadMaxUsos, SqlDbType.Int);
            AddParam(cmd, "@FechaInicio", model.FechaInicio.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@FechaFin", model.FechaFin.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> CuponesEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Cupones_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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
