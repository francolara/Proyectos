using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<int> NotificacionesContarNoLeidasAsync(int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Notificaciones_ContarNoLeidas", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        var value = await cmd.ExecuteScalarAsync();
        return value is null || value == DBNull.Value ? 0 : Convert.ToInt32(value);
    }

    public async Task<List<NotificacionNegocioItemViewModel>> NotificacionesListarAsync(int negocioId, int top = 15)
    {
        var list = new List<NotificacionNegocioItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Notificaciones_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Top", top, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new NotificacionNegocioItemViewModel
            {
                Id = dr.GetInt32(0),
                Tipo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Titulo = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
                Mensaje = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                UrlDestino = dr.IsDBNull(4) ? null : dr.GetString(4),
                FechaRegistroUtc = dr.GetDateTime(5)
            });
        }

        return list;
    }

    public async Task<bool> NotificacionesMarcarLeidaAsync(int negocioId, int notificacionId, string? userId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Notificaciones_MarcarLeida", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@NotificacionId", notificacionId, SqlDbType.Int);
        AddParam(cmd, "@UserId", string.IsNullOrWhiteSpace(userId) ? null : userId.Trim(), SqlDbType.NVarChar);
        var rows = await cmd.ExecuteNonQueryAsync();
        return rows > 0;
    }

    public async Task<int> NotificacionesMarcarTodasLeidasAsync(int negocioId, string? userId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Notificaciones_MarcarTodasLeidas", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@UserId", string.IsNullOrWhiteSpace(userId) ? null : userId.Trim(), SqlDbType.NVarChar);
        var value = await cmd.ExecuteScalarAsync();
        return value is null || value == DBNull.Value ? 0 : Convert.ToInt32(value);
    }
}
