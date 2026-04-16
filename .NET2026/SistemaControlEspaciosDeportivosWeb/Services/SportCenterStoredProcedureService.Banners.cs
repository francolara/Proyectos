using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<WebBannerPublicoViewModel>> HomeListarBannersPublicosAsync()
    {
        var list = new List<WebBannerPublicoViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_ListarBannersPublicos", cn) { CommandType = CommandType.StoredProcedure };
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new WebBannerPublicoViewModel
            {
                Id = dr.GetInt32(0),
                Titulo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Subtitulo = dr.IsDBNull(2) ? null : dr.GetString(2),
                Descripcion = dr.IsDBNull(3) ? null : dr.GetString(3),
                BotonTexto = dr.IsDBNull(4) ? null : dr.GetString(4),
                BotonUrl = dr.IsDBNull(5) ? null : dr.GetString(5),
                ImagenUrl = dr.IsDBNull(6) ? string.Empty : dr.GetString(6),
                ImagenUrlMobile = dr.IsDBNull(7) ? null : dr.GetString(7),
                Orden = dr.IsDBNull(8) ? 0 : dr.GetInt32(8)
            });
        }

        return list;
    }

    public async Task<WebBannerPublicoViewModel?> WebBannersObtenerFijoPorTipoAsync(int tipoBanner)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_WebBanners_ObtenerFijoPorTipo", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@TipoBanner", tipoBanner, SqlDbType.TinyInt);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
            return null;

        return new WebBannerPublicoViewModel
        {
            Id = dr.GetInt32(0),
            Titulo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
            Subtitulo = dr.IsDBNull(2) ? null : dr.GetString(2),
            Descripcion = dr.IsDBNull(3) ? null : dr.GetString(3),
            BotonTexto = dr.IsDBNull(4) ? null : dr.GetString(4),
            BotonUrl = dr.IsDBNull(5) ? null : dr.GetString(5),
            ImagenUrl = dr.IsDBNull(6) ? string.Empty : dr.GetString(6),
            ImagenUrlMobile = dr.IsDBNull(7) ? null : dr.GetString(7),
            Orden = dr.IsDBNull(8) ? 0 : dr.GetInt32(8)
        };
    }

    public async Task<List<BannerAdminItemViewModel>> BannersAdminListarAsync(bool? soloActivos = null)
    {
        var list = new List<BannerAdminItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_WebBanners_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@SoloActivos", soloActivos, SqlDbType.Bit);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new BannerAdminItemViewModel
            {
                Id = dr.GetInt32(0),
                Titulo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Subtitulo = dr.IsDBNull(2) ? null : dr.GetString(2),
                Descripcion = dr.IsDBNull(3) ? null : dr.GetString(3),
                BotonTexto = dr.IsDBNull(4) ? null : dr.GetString(4),
                BotonUrl = dr.IsDBNull(5) ? null : dr.GetString(5),
                ImagenUrl = dr.IsDBNull(6) ? string.Empty : dr.GetString(6),
                ImagenUrlMobile = dr.IsDBNull(7) ? null : dr.GetString(7),
                TipoBanner = dr.IsDBNull(8) ? 1 : Convert.ToInt32(dr.GetValue(8)),
                TipoBannerNombre = dr.IsDBNull(9) ? "Home" : dr.GetString(9),
                Orden = dr.IsDBNull(10) ? 0 : dr.GetInt32(10),
                Activo = !dr.IsDBNull(11) && ReadBool(dr, 11),
                FechaInicio = dr.IsDBNull(12) ? null : DateOnly.FromDateTime(dr.GetDateTime(12)),
                FechaFin = dr.IsDBNull(13) ? null : DateOnly.FromDateTime(dr.GetDateTime(13))
            });
        }

        return list;
    }

    public async Task<int> BannersAdminGuardarAsync(BannerAdminFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_WebBanners_Guardar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@Titulo", model.Titulo, SqlDbType.NVarChar);
        AddParam(cmd, "@Subtitulo", model.Subtitulo, SqlDbType.NVarChar);
        AddParam(cmd, "@Descripcion", model.Descripcion, SqlDbType.NVarChar);
        AddParam(cmd, "@BotonTexto", model.BotonTexto, SqlDbType.NVarChar);
        AddParam(cmd, "@BotonUrl", model.BotonUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@ImagenUrl", model.ImagenUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@ImagenUrlMobile", model.ImagenUrlMobile, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoBanner", model.TipoBanner, SqlDbType.TinyInt);
        AddParam(cmd, "@Orden", model.Orden, SqlDbType.Int);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@FechaInicio", model.FechaInicio?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaFin", model.FechaFin?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result ?? 0);
    }

    public async Task<bool> BannersAdminEliminarAsync(int id, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_WebBanners_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }
}
