using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<SolicitudPublicaItemViewModel>> SolicitudesPublicasListarAsync(int negocioId, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int? estado = null)
    {
        var list = new List<SolicitudPublicaItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_SolicitudesPublicas_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@Estado", estado, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SolicitudPublicaItemViewModel
            {
                Id = dr.GetInt32(0),
                CodigoSolicitud = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Sede = dr.GetString(2),
                Espacio = dr.GetString(3),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(4)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(5)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(6)),
                NombreSolicitante = dr.GetString(7),
                Telefono = dr.GetString(8),
                Correo = dr.IsDBNull(9) ? null : dr.GetString(9),
                Estado = dr.GetInt32(10),
                ReservaId = dr.IsDBNull(11) ? null : dr.GetInt32(11),
                FechaRegistro = dr.GetDateTime(12)
            });
        }
        return list;
    }

    public async Task<bool> SolicitudesPublicasActualizarEstadoAsync(SolicitudEstadoFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_SolicitudesPublicas_ActualizarEstado", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@Estado", model.Estado, SqlDbType.Int);
        AddParam(cmd, "@ComentarioGestion", model.ComentarioGestion, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<int> SolicitudesPublicasConvertirAReservaAsync(SolicitudConvertirFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_SolicitudesPublicas_ConvertirAReserva", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@Total", model.Total, SqlDbType.Decimal);
        AddParam(cmd, "@Adelanto", model.Adelanto, SqlDbType.Decimal);
        AddParam(cmd, "@EstadoReserva", model.EstadoReserva, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }
}
