using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<string> HomeSolicitarAltaClubAsync(AltaClubSolicitudFormViewModel model)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_SolicitarAltaClub", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NombreContacto", model.NombreContacto, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@RelacionClub", model.RelacionClub, SqlDbType.NVarChar);
        AddParam(cmd, "@NombreClub", model.NombreClub, SqlDbType.NVarChar);
        AddParam(cmd, "@Pais", model.Pais, SqlDbType.NVarChar);
        AddParam(cmd, "@ProvinciaEstado", model.ProvinciaEstado, SqlDbType.NVarChar);
        AddParam(cmd, "@Ciudad", model.Ciudad, SqlDbType.NVarChar);
        AddParam(cmd, "@Direccion", model.Direccion, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return result?.ToString() ?? string.Empty;
    }

    public async Task<List<AltaClubItemViewModel>> AltasClubesListarAsync(int? estado = null)
    {
        var list = new List<AltaClubItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_AltasClubes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Estado", estado, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new AltaClubItemViewModel
            {
                Id = dr.GetInt32(0),
                CodigoSolicitud = dr.GetString(1),
                NombreContacto = dr.GetString(2),
                Telefono = dr.GetString(3),
                Correo = dr.GetString(4),
                RelacionClub = dr.GetString(5),
                NombreClub = dr.GetString(6),
                Pais = dr.GetString(7),
                ProvinciaEstado = dr.GetString(8),
                Ciudad = dr.GetString(9),
                Direccion = dr.GetString(10),
                Estado = dr.GetInt32(11),
                ComentarioGestion = dr.IsDBNull(12) ? null : dr.GetString(12),
                NegocioId = dr.IsDBNull(13) ? null : dr.GetInt32(13),
                SedeId = dr.IsDBNull(14) ? null : dr.GetInt32(14),
                FechaRegistro = dr.GetDateTime(15),
                FechaGestion = dr.IsDBNull(16) ? null : dr.GetDateTime(16)
            });
        }

        return list;
    }

    public async Task<string> HomeRegistrarClubConPruebaAsync(AltaClubSolicitudFormViewModel model, string usuarioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_RegistrarClubConPrueba", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@NombreContacto", model.NombreContacto, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@RelacionClub", model.RelacionClub, SqlDbType.NVarChar);
        AddParam(cmd, "@NombreClub", model.NombreClub, SqlDbType.NVarChar);
        AddParam(cmd, "@Pais", model.Pais, SqlDbType.NVarChar);
        AddParam(cmd, "@ProvinciaEstado", model.ProvinciaEstado, SqlDbType.NVarChar);
        AddParam(cmd, "@Ciudad", model.Ciudad, SqlDbType.NVarChar);
        AddParam(cmd, "@Direccion", model.Direccion, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return result?.ToString() ?? string.Empty;
    }

    public async Task<bool> AltasClubesAprobarAsync(int id, string usuario, string? comentarioGestion = null, int diasPrueba = 30)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_AltasClubes_Aprobar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            AddParam(cmd, "@ComentarioGestion", comentarioGestion, SqlDbType.NVarChar);
            AddParam(cmd, "@DiasPrueba", diasPrueba <= 0 ? 30 : diasPrueba, SqlDbType.Int);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> AltasClubesRechazarAsync(int id, string usuario, string? comentarioGestion = null)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_AltasClubes_Rechazar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            AddParam(cmd, "@ComentarioGestion", comentarioGestion, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }
}
