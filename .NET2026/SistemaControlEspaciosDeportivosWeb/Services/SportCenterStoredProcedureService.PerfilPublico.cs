using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<UsuarioPublicoPerfilViewModel?> UsuariosPublicosObtenerPerfilAsync(string usuarioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_ObtenerPerfil", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new UsuarioPublicoPerfilViewModel
        {
            Id = dr.GetInt32(0),
            UsuarioId = dr.GetString(1),
            TipoDocumento = dr.IsDBNull(2) ? "0" : dr.GetString(2),
            NumeroDocumento = dr.IsDBNull(3) ? null : dr.GetString(3),
            Nombres = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
            Apellidos = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
            NombreEquipo = dr.IsDBNull(6) ? null : dr.GetString(6),
            Telefono = dr.IsDBNull(7) ? null : dr.GetString(7),
            Correo = dr.IsDBNull(8) ? null : dr.GetString(8),
            FechaNacimiento = dr.IsDBNull(9) ? null : DateOnly.FromDateTime(dr.GetDateTime(9)),
            CodigoUbigeo = dr.IsDBNull(10) ? null : dr.GetString(10),
            CodigoDepartamento = dr.IsDBNull(11) ? null : dr.GetString(11),
            CodigoProvincia = dr.IsDBNull(12) ? null : dr.GetString(12)
        };
    }

    public async Task<int> UsuariosPublicosGuardarPerfilAsync(UsuarioPublicoPerfilViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_GuardarPerfil", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", model.UsuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoDocumento", model.TipoDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@NumeroDocumento", model.NumeroDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@Nombres", model.Nombres, SqlDbType.NVarChar);
        AddParam(cmd, "@Apellidos", model.Apellidos, SqlDbType.NVarChar);
        AddParam(cmd, "@NombreEquipo", model.NombreEquipo, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@FechaNacimiento", model.FechaNacimiento?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result);
    }

    public async Task<List<UsuarioPublicoReservaItemViewModel>> UsuariosPublicosReservasListarAsync(string usuarioId, int top = 200)
    {
        var list = new List<UsuarioPublicoReservaItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_ReservasListar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@Top", top <= 0 ? 200 : top, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new UsuarioPublicoReservaItemViewModel
            {
                ReservaId = dr.GetInt32(0),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(1)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(2)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(3)),
                EstadoTexto = dr.GetString(4),
                Total = dr.GetDecimal(5),
                Adelanto = dr.GetDecimal(6),
                SaldoPendiente = dr.GetDecimal(7),
                NegocioNombre = dr.GetString(8),
                SedeNombre = dr.GetString(9),
                EspacioNombre = dr.GetString(10),
                SedeDireccion = dr.IsDBNull(11) ? null : dr.GetString(11),
                SedeTelefono = dr.IsDBNull(12) ? null : dr.GetString(12),
                SedeWhatsapp = dr.IsDBNull(13) ? null : dr.GetString(13)
            });
        }

        return list;
    }
}
