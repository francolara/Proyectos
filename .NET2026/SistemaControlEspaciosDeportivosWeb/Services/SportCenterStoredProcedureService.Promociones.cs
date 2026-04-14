using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<(List<PromocionItemViewModel> Promociones, int TotalRegistros)> PromocionesListarAsync(
        int negocioId,
        int? sedeId = null,
        DateOnly? fechaDesde = null,
        DateOnly? fechaHasta = null,
        bool? soloActivos = true,
        int pagina = 1,
        int tamanoPagina = 20)
    {
        var list = new List<PromocionItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Promociones_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SoloActivos", soloActivos, SqlDbType.Bit);
        AddParam(cmd, "@Pagina", pagina, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoPagina, SqlDbType.Int);
        var totalRegistrosParam = new SqlParameter("@TotalRegistros", SqlDbType.Int) { Direction = ParameterDirection.Output };
        cmd.Parameters.Add(totalRegistrosParam);
        await using var dr = await cmd.ExecuteReaderAsync();

        while (await dr.ReadAsync())
        {
            list.Add(new PromocionItemViewModel
            {
                Id = dr.GetInt32(0),
                Nombre = dr.GetString(1),
                Sede = dr.GetString(2),
                Espacio = dr.GetString(3),
                FechaInicio = DateOnly.FromDateTime(dr.GetDateTime(4)),
                FechaFin = DateOnly.FromDateTime(dr.GetDateTime(5)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(6)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(7)),
                PorcentajeDescuento = dr.GetDecimal(8),
                Activo = dr.GetBoolean(9)
            });
        }

        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        return (list, totalRegistros);
    }

    public async Task<PromocionFormViewModel?> PromocionesObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Promociones_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new PromocionFormViewModel
        {
            Id = dr.GetInt32(0),
            SedeId = dr.IsDBNull(1) ? null : dr.GetInt32(1),
            EspacioDeportivoId = dr.IsDBNull(2) ? null : dr.GetInt32(2),
            Nombre = dr.GetString(3),
            FechaInicio = DateOnly.FromDateTime(dr.GetDateTime(4)),
            FechaFin = DateOnly.FromDateTime(dr.GetDateTime(5)),
            HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(6)),
            HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(7)),
            PorcentajeDescuento = dr.GetDecimal(8),
            Activo = dr.GetBoolean(9),
            NegocioId = negocioId
        };
    }

    public async Task<int> PromocionesCrearAsync(PromocionFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Promociones_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", model.SedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@FechaInicio", model.FechaInicio.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaFin", model.FechaFin.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", model.HoraInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", model.HoraFin.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@PorcentajeDescuento", model.PorcentajeDescuento, SqlDbType.Decimal);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> PromocionesActualizarAsync(PromocionFormViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Promociones_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@SedeId", model.SedeId, SqlDbType.Int);
            AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
            AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
            AddParam(cmd, "@FechaInicio", model.FechaInicio.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@FechaFin", model.FechaFin.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@HoraInicio", model.HoraInicio.ToTimeSpan(), SqlDbType.Time);
            AddParam(cmd, "@HoraFin", model.HoraFin.ToTimeSpan(), SqlDbType.Time);
            AddParam(cmd, "@PorcentajeDescuento", model.PorcentajeDescuento, SqlDbType.Decimal);
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

    public async Task<bool> PromocionesEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Promociones_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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
