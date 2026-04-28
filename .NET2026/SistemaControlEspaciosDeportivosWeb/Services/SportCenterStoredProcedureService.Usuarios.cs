using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<SelectListItem>> UsuariosNegocioRolesListarAsync()
        => await ComboAsync("Sp_UsuariosNegocio_RolesListar");

    public async Task<List<UsuarioNegocioItemViewModel>> UsuariosNegocioListarAsync(int negocioId, int? sedeId = null)
    {
        var list = new List<UsuarioNegocioItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosNegocio_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new UsuarioNegocioItemViewModel
            {
                UsuarioNegocioId = dr.GetInt32(0),
                UsuarioId = dr.GetString(1),
                Nombres = dr.GetString(2),
                Apellidos = dr.GetString(3),
                Correo = dr.GetString(4),
                RolNegocio = dr.GetInt32(5),
                Activo = dr.GetBoolean(6),
                SedeId = dr.IsDBNull(7) ? null : dr.GetInt32(7),
                SedeNombre = dr.IsDBNull(8) ? string.Empty : dr.GetString(8)
            });
        }

        return list;
    }

    public async Task<bool> UsuariosNegocioAsignarPorCorreoAsync(int negocioId, string correo, int rolNegocio, int? sedeId, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosNegocio_AsignarPorCorreo", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Correo", correo, SqlDbType.NVarChar);
        AddParam(cmd, "@RolNegocio", rolNegocio, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }

    public async Task<bool> UsuariosNegocioActualizarRolAsync(int negocioId, int usuarioNegocioId, int rolNegocio, int? sedeId, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_UsuariosNegocio_ActualizarRol", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@UsuarioNegocioId", usuarioNegocioId, SqlDbType.Int);
            AddParam(cmd, "@RolNegocio", rolNegocio, SqlDbType.Int);
            AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> UsuariosNegocioDesactivarAsync(int negocioId, int usuarioNegocioId, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_UsuariosNegocio_Desactivar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@UsuarioNegocioId", usuarioNegocioId, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<List<UsuarioNegocioPermisoModuloViewModel>> UsuariosNegocioPermisosListarAsync(int negocioId, int usuarioNegocioId)
    {
        var list = new List<UsuarioNegocioPermisoModuloViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosNegocio_PermisosListar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@UsuarioNegocioId", usuarioNegocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new UsuarioNegocioPermisoModuloViewModel
            {
                ModuloSistemaId = dr.GetInt32(0),
                ModuloCodigo = dr.GetString(1),
                ModuloNombre = dr.GetString(2),
                PuedeVer = dr.GetBoolean(3),
                PuedeCrear = dr.GetBoolean(4),
                PuedeEditar = dr.GetBoolean(5),
                PuedeEliminar = dr.GetBoolean(6)
            });
        }
        return list;
    }

    public async Task<bool> UsuariosNegocioPermisoGuardarAsync(int negocioId, int usuarioNegocioId, UsuarioNegocioPermisoModuloViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosNegocio_PermisoGuardar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@UsuarioNegocioId", usuarioNegocioId, SqlDbType.Int);
        AddParam(cmd, "@ModuloSistemaId", model.ModuloSistemaId, SqlDbType.Int);
        AddParam(cmd, "@PuedeVer", model.PuedeVer, SqlDbType.Bit);
        AddParam(cmd, "@PuedeCrear", model.PuedeCrear, SqlDbType.Bit);
        AddParam(cmd, "@PuedeEditar", model.PuedeEditar, SqlDbType.Bit);
        AddParam(cmd, "@PuedeEliminar", model.PuedeEliminar, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }
}
