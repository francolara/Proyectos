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
            var reservaId = dr.IsDBNull(11) ? (int?)null : dr.GetInt32(11);
            int? numeroReservaPorNegocio = reservaId.HasValue
                ? ReadRequiredInt32(dr, 12, "NumeroReservaPorNegocio")
                : null;

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
                ReservaId = reservaId,
                NumeroReservaPorNegocio = numeroReservaPorNegocio,
                FechaRegistro = dr.GetDateTime(13)
            });
        }
        return list;
    }

    public async Task<bool> SolicitudesPublicasActualizarEstadoAsync(SolicitudEstadoFormViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_SolicitudesPublicas_ActualizarEstado", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@Estado", model.Estado, SqlDbType.Int);
            AddParam(cmd, "@ComentarioGestion", model.ComentarioGestion, SqlDbType.NVarChar);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<(int ReservaId, int NumeroPorNegocio)> SolicitudesPublicasConvertirAReservaAsync(SolicitudConvertirFormViewModel model, string usuario)
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
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync() || dr.FieldCount < 2 || dr.IsDBNull(1))
            throw new InvalidOperationException("El procedimiento no devolvio el correlativo visible de la reserva creada.");

        return (dr.GetInt32(0), dr.GetInt32(1));
    }
}
