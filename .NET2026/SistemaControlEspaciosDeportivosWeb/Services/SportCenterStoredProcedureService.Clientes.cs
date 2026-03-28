using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<ClienteItemViewModel>> ClientesListarAsync(int negocioId)
    {
        var list = new List<ClienteItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Clientes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ClienteItemViewModel
            {
                Id = dr.GetInt32(0),
                NombresORazonSocial = dr.GetString(1),
                TipoDocumento = dr.GetString(2),
                NumeroDocumento = dr.GetString(3),
                Telefono = dr.IsDBNull(4) ? null : dr.GetString(4),
                Correo = dr.IsDBNull(5) ? null : dr.GetString(5),
                Activo = dr.GetBoolean(6)
            });
        }
        return list;
    }

    public async Task<ClienteFormViewModel?> ClientesObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Clientes_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        return new ClienteFormViewModel
        {
            Id = dr.GetInt32(0),
            NombresORazonSocial = dr.GetString(1),
            TipoDocumento = dr.GetString(2),
            NumeroDocumento = dr.GetString(3),
            Telefono = dr.IsDBNull(4) ? null : dr.GetString(4),
            Correo = dr.IsDBNull(5) ? null : dr.GetString(5),
            DireccionFiscal = dr.IsDBNull(6) ? null : dr.GetString(6),
            Activo = dr.GetBoolean(7),
            NegocioId = negocioId
        };
    }

    public async Task<int> ClientesCrearAsync(ClienteFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Clientes_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@NombresORazonSocial", model.NombresORazonSocial, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoDocumento", model.TipoDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@NumeroDocumento", model.NumeroDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@DireccionFiscal", model.DireccionFiscal, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> ClientesActualizarAsync(ClienteFormViewModel model, string usuario)
    {
        var clienteActual = await ClientesObtenerAsync(model.NegocioId, model.Id);
        if (clienteActual is null) return false;

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Clientes_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@NombresORazonSocial", model.NombresORazonSocial, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoDocumento", model.TipoDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@NumeroDocumento", model.NumeroDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@DireccionFiscal", model.DireccionFiscal, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }

    public async Task<bool> ClientesEliminarAsync(int negocioId, int id, string usuario)
    {
        var clienteActual = await ClientesObtenerAsync(negocioId, id);
        if (clienteActual is null) return false;

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Clientes_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }
}
