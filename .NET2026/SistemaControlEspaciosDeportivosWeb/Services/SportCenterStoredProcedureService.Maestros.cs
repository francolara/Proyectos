using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<MonedaMaestroItemViewModel>> MaestrosMonedasListarAsync(int negocioId)
    {
        var list = new List<MonedaMaestroItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_Monedas_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new MonedaMaestroItemViewModel
            {
                Id = dr.GetInt32(0),
                MonedaSuperId = dr.GetInt32(1),
                Codigo = dr.GetString(2),
                Nombre = dr.GetString(3),
                Simbolo = dr.IsDBNull(4) ? null : dr.GetString(4),
                Activo = dr.GetBoolean(5)
            });
        }
        return list;
    }

    public async Task<List<SelectListItem>> MaestrosMonedasSuperListarAsync()
        => await ComboAsync("Sp_Maestros_MonedasSuper_Listar");

    public async Task<int> MaestrosMonedasCrearAsync(int negocioId, int monedaSuperId, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_Monedas_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@MonedaSuperId", monedaSuperId, SqlDbType.Int);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> MaestrosMonedasActualizarAsync(int negocioId, int id, bool activo, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_Monedas_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> MaestrosMonedasEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_Monedas_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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

    public async Task<List<MaestroCatalogoItemViewModel>> MaestrosTiposSueloListarAsync(int negocioId)
    {
        var list = new List<MaestroCatalogoItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_TiposSuelo_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new MaestroCatalogoItemViewModel
            {
                Id = dr.GetInt32(0),
                SuperId = dr.IsDBNull(1) ? null : dr.GetInt32(1),
                Codigo = dr.IsDBNull(2) ? null : dr.GetString(2),
                Nombre = dr.GetString(3),
                Activo = dr.GetBoolean(4)
            });
        }
        return list;
    }

    public async Task<List<SelectListItem>> MaestrosTiposSueloSuperListarAsync()
        => await ComboAsync("Sp_Maestros_TiposSueloSuper_Listar");

    public async Task<int> MaestrosTiposSueloCrearAsync(int negocioId, int tipoSueloSuperId, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_TiposSuelo_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@TipoSueloSuperId", tipoSueloSuperId, SqlDbType.Int);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> MaestrosTiposSueloActualizarAsync(int negocioId, int id, bool activo, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_TiposSuelo_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> MaestrosTiposSueloEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_TiposSuelo_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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

    public async Task<List<MaestroCatalogoItemViewModel>> MaestrosTiposDeporteListarAsync(int negocioId)
    {
        var list = new List<MaestroCatalogoItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_TiposDeporte_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new MaestroCatalogoItemViewModel
            {
                Id = dr.GetInt32(0),
                SuperId = dr.IsDBNull(1) ? null : dr.GetInt32(1),
                Codigo = dr.IsDBNull(2) ? null : dr.GetString(2),
                Nombre = dr.GetString(3),
                Activo = dr.GetBoolean(4)
            });
        }
        return list;
    }

    public async Task<List<SelectListItem>> MaestrosTiposDeporteSuperListarAsync()
        => await ComboAsync("Sp_Maestros_TiposDeporteSuper_Listar");

    public async Task<int> MaestrosTiposDeporteCrearAsync(int negocioId, int tipoDeporteSuperId, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_TiposDeporte_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@TipoDeporteSuperId", tipoDeporteSuperId, SqlDbType.Int);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> MaestrosTiposDeporteActualizarAsync(int negocioId, int id, bool activo, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_TiposDeporte_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> MaestrosTiposDeporteEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_TiposDeporte_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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

    public async Task<List<MaestroCatalogoItemViewModel>> MaestrosFormasPagoListarAsync(int negocioId)
    {
        var list = new List<MaestroCatalogoItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_FormasPago_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new MaestroCatalogoItemViewModel
            {
                Id = dr.GetInt32(0),
                Nombre = dr.GetString(1),
                Activo = dr.GetBoolean(2)
            });
        }
        return list;
    }

    public async Task<int> MaestrosFormasPagoCrearAsync(int negocioId, string nombre, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_FormasPago_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Nombre", nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> MaestrosFormasPagoActualizarAsync(int negocioId, int id, string nombre, bool activo, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_FormasPago_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Nombre", nombre, SqlDbType.NVarChar);
            AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> MaestrosFormasPagoEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_FormasPago_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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

    public Task<List<SelectListItem>> MaestrosTiposDocumentoComprobanteSuperListarAsync()
        => ComboAsync("Sp_Maestros_TiposDocumentoComprobanteSuper_Listar");

    public async Task<List<TipoDocumentoComprobanteNegocioItemViewModel>> MaestrosTiposDocumentoComprobanteListarAsync(int negocioId)
    {
        var list = new List<TipoDocumentoComprobanteNegocioItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_TiposDocumentoComprobante_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new TipoDocumentoComprobanteNegocioItemViewModel
            {
                Id = dr.GetInt32(0),
                CodigoSunat = dr.GetString(1),
                Nombre = dr.GetString(2),
                Tributario = dr.GetBoolean(3),
                HabilitadoSuper = dr.GetBoolean(4),
                Activo = dr.GetBoolean(5)
            });
        }
        return list;
    }

    public async Task<int> MaestrosTiposDocumentoComprobanteCrearAsync(int negocioId, string codigoSunat, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Maestros_TiposDocumentoComprobante_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@CodigoSunat", codigoSunat, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> MaestrosTiposDocumentoComprobanteActualizarAsync(int negocioId, int id, bool activo, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_TiposDocumentoComprobante_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> MaestrosTiposDocumentoComprobanteEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Maestros_TiposDocumentoComprobante_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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
