using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<(List<ClienteItemViewModel> Clientes, int TotalRegistros, int TotalActivos, int TotalInactivos)> ClientesListarAsync(int negocioId, bool? activo = null, string? buscar = null, int pagina = 1, int tamanoPagina = 20)
    {
        var list = new List<ClienteItemViewModel>();
        var paginaNormalizada = pagina < 1 ? 1 : pagina;
        var tamanoNormalizado = tamanoPagina < 1 ? 20 : tamanoPagina;
        var textoBusqueda = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Clientes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Buscar", textoBusqueda, SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", paginaNormalizada, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoNormalizado, SqlDbType.Int);
        var totalRegistrosParam = cmd.Parameters.Add("@TotalRegistros", SqlDbType.Int);
        totalRegistrosParam.Direction = ParameterDirection.Output;
        var totalActivosParam = cmd.Parameters.Add("@TotalActivos", SqlDbType.Int);
        totalActivosParam.Direction = ParameterDirection.Output;
        var totalInactivosParam = cmd.Parameters.Add("@TotalInactivos", SqlDbType.Int);
        totalInactivosParam.Direction = ParameterDirection.Output;
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ClienteItemViewModel
            {
                Id = dr.GetInt32(0),
                NombresORazonSocial = dr.GetString(1),
                NombreEquipo = dr.IsDBNull(2) ? null : dr.GetString(2),
                TipoDocumento = dr.GetString(3),
                NumeroDocumento = dr.GetString(4),
                Telefono = dr.IsDBNull(5) ? null : dr.GetString(5),
                Correo = dr.IsDBNull(6) ? null : dr.GetString(6),
                Activo = dr.GetBoolean(7)
            });
        }
        await dr.CloseAsync();

        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        var totalActivos = totalActivosParam.Value is int activos ? activos : 0;
        var totalInactivos = totalInactivosParam.Value is int inactivos ? inactivos : 0;
        return (list, totalRegistros, totalActivos, totalInactivos);
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
        if (dr.FieldCount < 12)
        {
            return new ClienteFormViewModel
            {
                Id = dr.GetInt32(0),
                NombresORazonSocial = dr.GetString(1),
                Nombres = dr.GetString(1),
                NombreEquipo = dr.IsDBNull(2) ? null : dr.GetString(2),
                TipoDocumento = dr.GetString(3),
                NumeroDocumento = dr.GetString(4),
                Telefono = dr.IsDBNull(5) ? null : dr.GetString(5),
                Correo = dr.IsDBNull(6) ? null : dr.GetString(6),
                DireccionFiscal = dr.IsDBNull(7) ? null : dr.GetString(7),
                CodigoUbigeo = dr.IsDBNull(8) ? null : dr.GetString(8),
                Activo = dr.GetBoolean(9),
                NegocioId = negocioId
            };
        }

        return new ClienteFormViewModel
        {
            Id = dr.GetInt32(0),
            NombresORazonSocial = dr.GetString(1),
            Nombres = dr.IsDBNull(2) ? null : dr.GetString(2),
            Apellidos = dr.IsDBNull(3) ? null : dr.GetString(3),
            NombreEquipo = dr.IsDBNull(4) ? null : dr.GetString(4),
            TipoDocumento = dr.GetString(5),
            NumeroDocumento = dr.GetString(6),
            Telefono = dr.IsDBNull(7) ? null : dr.GetString(7),
            Correo = dr.IsDBNull(8) ? null : dr.GetString(8),
            DireccionFiscal = dr.IsDBNull(9) ? null : dr.GetString(9),
            CodigoUbigeo = dr.IsDBNull(10) ? null : dr.GetString(10),
            Activo = dr.GetBoolean(11),
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
        AddParam(cmd, "@Nombres", model.Nombres, SqlDbType.NVarChar);
        AddParam(cmd, "@Apellidos", model.Apellidos, SqlDbType.NVarChar);
        AddParam(cmd, "@NombreEquipo", model.NombreEquipo, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoDocumento", model.TipoDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@NumeroDocumento", model.NumeroDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@DireccionFiscal", model.DireccionFiscal, SqlDbType.NVarChar);
        AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> ClientesActualizarAsync(ClienteFormViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Clientes_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@NombresORazonSocial", model.NombresORazonSocial, SqlDbType.NVarChar);
            AddParam(cmd, "@Nombres", model.Nombres, SqlDbType.NVarChar);
            AddParam(cmd, "@Apellidos", model.Apellidos, SqlDbType.NVarChar);
            AddParam(cmd, "@NombreEquipo", model.NombreEquipo, SqlDbType.NVarChar);
            AddParam(cmd, "@TipoDocumento", model.TipoDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@NumeroDocumento", model.NumeroDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
            AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
            AddParam(cmd, "@DireccionFiscal", model.DireccionFiscal, SqlDbType.NVarChar);
            AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
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

    public async Task<bool> ClientesEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Clientes_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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
