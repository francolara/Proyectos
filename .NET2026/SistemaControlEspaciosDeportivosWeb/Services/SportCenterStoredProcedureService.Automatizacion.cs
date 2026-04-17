using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<ReservaRecordatorioPendienteViewModel>> ReservasRecordatoriosPendientesAsync(DateTime fechaHoraActual)
    {
        var list = new List<ReservaRecordatorioPendienteViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_RecordatoriosPendientes", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@FechaHoraActual", fechaHoraActual, SqlDbType.DateTime2);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReservaRecordatorioPendienteViewModel
            {
                ReservaId = dr.GetInt32(0),
                NegocioId = dr.GetInt32(1),
                Cliente = dr.GetString(2),
                Correo = dr.GetString(3),
                Sede = dr.GetString(4),
                Espacio = dr.GetString(5),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(6)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(7)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(8)),
                CorreoNotificacion = dr.IsDBNull(9) ? null : dr.GetString(9),
                WhatsappContacto = dr.IsDBNull(10) ? null : dr.GetString(10)
            });
        }

        return list;
    }

    public async Task<bool> ReservasMarcarRecordatorioEnviadoAsync(int negocioId, int reservaId, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Reservas_MarcarRecordatorioEnviado", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@ReservaId", reservaId, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<int> ReservasAutoNoShowAsync(DateTime fechaHoraActual, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_AutoNoShow", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@FechaHoraActual", fechaHoraActual, SqlDbType.DateTime2);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result ?? 0);
    }

    public async Task<int> ReservasAutoCancelarNoConfirmadasAsync(DateTime fechaHoraActual, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_AutoCancelarNoConfirmadas", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@FechaHoraActual", fechaHoraActual, SqlDbType.DateTime2);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result ?? 0);
    }
}
