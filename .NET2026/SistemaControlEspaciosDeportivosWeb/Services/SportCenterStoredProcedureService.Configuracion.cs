using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<ConfiguracionClubViewModel?> ConfiguracionClubObtenerAsync(int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_ConfiguracionClub_Obtener", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new ConfiguracionClubViewModel
        {
            Id = dr.GetInt32(0),
            NegocioId = dr.GetInt32(0),
            NombreComercial = dr.GetString(1),
            RazonSocial = dr.IsDBNull(2) ? null : dr.GetString(2),
            TipoDocumento = dr.IsDBNull(3) ? "DNI" : dr.GetString(3),
            NumeroDocumento = dr.IsDBNull(4) ? null : dr.GetString(4),
            DireccionFiscal = dr.IsDBNull(5) ? null : dr.GetString(5),
            MonedaId = dr.IsDBNull(6) ? 1 : dr.GetInt32(6)
        };
    }

    public async Task<bool> ConfiguracionClubActualizarAsync(ConfiguracionClubViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_ConfiguracionClub_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@NombreComercial", model.NombreComercial, SqlDbType.NVarChar);
            AddParam(cmd, "@RazonSocial", model.RazonSocial, SqlDbType.NVarChar);
            AddParam(cmd, "@TipoDocumentoFiscal", model.TipoDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@NumeroDocumentoFiscal", model.NumeroDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@DireccionFiscal", model.DireccionFiscal, SqlDbType.NVarChar);
            AddParam(cmd, "@MonedaId", model.MonedaId, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<List<SelectListItem>> ConfiguracionClubComboMonedasAsync(int negocioId)
    {
        var list = new List<SelectListItem>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Combos_Monedas", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SelectListItem(dr.GetString(1), dr.GetInt32(0).ToString()));
        }
        return list;
    }
}
