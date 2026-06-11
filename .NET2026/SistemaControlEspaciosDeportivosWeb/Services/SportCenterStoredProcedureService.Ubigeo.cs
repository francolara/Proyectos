using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<SelectListItem>> UbigeoDepartamentosListarAsync()
    {
        var list = new List<SelectListItem>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Ubigeo_Departamentos_Listar", cn) { CommandType = CommandType.StoredProcedure };
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SelectListItem(dr.GetString(1), dr.GetString(0)));
        }
        return list;
    }

    public async Task<List<SelectListItem>> UbigeoProvinciasListarAsync(string codigoDepartamento)
    {
        var list = new List<SelectListItem>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Ubigeo_Provincias_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@CodigoDepartamento", codigoDepartamento, SqlDbType.Char);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SelectListItem(dr.GetString(1), dr.GetString(0)));
        }
        return list;
    }

    public async Task<List<SelectListItem>> UbigeoDistritosListarAsync(string codigoProvincia)
    {
        var list = new List<SelectListItem>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Ubigeo_Distritos_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@CodigoProvincia", codigoProvincia, SqlDbType.Char);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SelectListItem(dr.GetString(1), dr.GetString(0)));
        }
        return list;
    }

    public async Task<List<SelectListItem>> UbigeoZonasListarAsync(string? codigoDepartamento = null, string? codigoProvincia = null)
    {
        return await ComboAsync(
            "Sp_Ubigeo_Zonas_Listar",
            ("@CodigoDepartamento", string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim(), SqlDbType.Char),
            ("@CodigoProvincia", string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim(), SqlDbType.Char));
    }

    public async Task<UbigeoLookupViewModel?> UbigeoObtenerPorCodigoAsync(string codigoUbigeo)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Ubigeo_ObtenerPorCodigo", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@CodigoUbigeo", codigoUbigeo, SqlDbType.Char);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new UbigeoLookupViewModel
        {
            CodigoUbigeo = dr.GetString(0),
            CodigoDepartamento = dr.GetString(1),
            CodigoProvincia = dr.GetString(2),
            Departamento = dr.GetString(3),
            Provincia = dr.GetString(4),
            Distrito = dr.GetString(5),
            Zona = dr.FieldCount > 6 && !dr.IsDBNull(6) ? dr.GetString(6) : null
        };
    }
}
