using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<(List<EspacioResenaAdminItemViewModel> Resenas, int TotalRegistros, int TotalVisibles, int TotalOcultas, int TotalRespondidas)> EspaciosResenasListarAsync(int negocioId, int espacioDeportivoId, int pagina = 1, int tamanoPagina = 4)
    {
        var list = new List<EspacioResenaAdminItemViewModel>();
        var paginaNormalizada = pagina < 1 ? 1 : pagina;
        var tamanoNormalizado = tamanoPagina < 1 ? 4 : tamanoPagina;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_ResenasListar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Pagina", paginaNormalizada, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoNormalizado, SqlDbType.Int);
        var totalRegistrosParam = cmd.Parameters.Add("@TotalRegistros", SqlDbType.Int);
        totalRegistrosParam.Direction = ParameterDirection.Output;
        var totalVisiblesParam = cmd.Parameters.Add("@TotalVisibles", SqlDbType.Int);
        totalVisiblesParam.Direction = ParameterDirection.Output;
        var totalOcultasParam = cmd.Parameters.Add("@TotalOcultas", SqlDbType.Int);
        totalOcultasParam.Direction = ParameterDirection.Output;
        var totalRespondidasParam = cmd.Parameters.Add("@TotalRespondidas", SqlDbType.Int);
        totalRespondidasParam.Direction = ParameterDirection.Output;

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new EspacioResenaAdminItemViewModel
            {
                ResenaId = dr.GetInt32(0),
                ReservaId = dr.GetInt32(1),
                EspacioDeportivoId = dr.GetInt32(2),
                EspacioNombre = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                SedeId = dr.GetInt32(4),
                SedeNombre = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
                AliasPublico = dr.IsDBNull(6) ? string.Empty : dr.GetString(6),
                Comentario = dr.IsDBNull(7) ? string.Empty : dr.GetString(7),
                Respuesta = dr.IsDBNull(8) ? null : dr.GetString(8),
                Activo = !dr.IsDBNull(9) && ReadBool(dr, 9),
                FechaCreacion = dr.GetDateTime(10),
                ReservaFecha = DateOnly.FromDateTime(dr.GetDateTime(11)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(12)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(13))
            });
        }

        await dr.CloseAsync();

        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        var totalVisibles = totalVisiblesParam.Value is int visibles ? visibles : 0;
        var totalOcultas = totalOcultasParam.Value is int ocultas ? ocultas : 0;
        var totalRespondidas = totalRespondidasParam.Value is int respondidas ? respondidas : 0;
        return (list, totalRegistros, totalVisibles, totalOcultas, totalRespondidas);
    }

    public async Task<bool> EspaciosResenaGestionarAsync(EspacioResenaGestionViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_ResenaGestionar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@ResenaId", model.ResenaId, SqlDbType.Int);
        AddParam(cmd, "@Respuesta", model.Respuesta, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        await cmd.ExecuteNonQueryAsync();
        return true;
    }
}
