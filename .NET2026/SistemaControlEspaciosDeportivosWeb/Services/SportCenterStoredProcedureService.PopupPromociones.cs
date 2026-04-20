using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<PopupPromocionPublicoViewModel>> HomeListarPopupPromocionesActivasAsync()
    {
        var list = new List<PopupPromocionPublicoViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_ListarPopupPromocionesActivas", cn) { CommandType = CommandType.StoredProcedure };
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new PopupPromocionPublicoViewModel
            {
                IdPopupPromocion = dr.GetInt32(0),
                Titulo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Subtitulo = dr.IsDBNull(2) ? null : dr.GetString(2),
                Descripcion = dr.IsDBNull(3) ? null : dr.GetString(3),
                ImagenUrl = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
                TextoBoton = dr.IsDBNull(5) ? null : dr.GetString(5),
                UrlBoton = dr.IsDBNull(6) ? null : dr.GetString(6),
                UrlImagen = dr.IsDBNull(7) ? null : dr.GetString(7),
                Orden = dr.IsDBNull(8) ? 0 : dr.GetInt32(8),
                AbrirNuevaPestana = !dr.IsDBNull(9) && ReadBool(dr, 9),
                Orientacion = dr.IsDBNull(10) ? PopupPromocionPublicoViewModel.OrientacionVertical : dr.GetString(10)
            });
        }

        return list;
    }

    public async Task<List<PopupPromocionAdminItemViewModel>> PopupPromocionesAdminListarAsync(bool? soloActivos = null)
    {
        var list = new List<PopupPromocionAdminItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_PopupPromociones_ListarAdmin", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@SoloActivos", soloActivos, SqlDbType.Bit);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new PopupPromocionAdminItemViewModel
            {
                IdPopupPromocion = dr.GetInt32(0),
                Titulo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Subtitulo = dr.IsDBNull(2) ? null : dr.GetString(2),
                Descripcion = dr.IsDBNull(3) ? null : dr.GetString(3),
                ImagenUrl = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
                TextoBoton = dr.IsDBNull(5) ? null : dr.GetString(5),
                UrlBoton = dr.IsDBNull(6) ? null : dr.GetString(6),
                UrlImagen = dr.IsDBNull(7) ? null : dr.GetString(7),
                Orden = dr.IsDBNull(8) ? 0 : dr.GetInt32(8),
                Activo = !dr.IsDBNull(9) && ReadBool(dr, 9),
                FechaInicio = dr.IsDBNull(10) ? null : DateOnly.FromDateTime(dr.GetDateTime(10)),
                FechaFin = dr.IsDBNull(11) ? null : DateOnly.FromDateTime(dr.GetDateTime(11)),
                AbrirNuevaPestana = !dr.IsDBNull(12) && ReadBool(dr, 12),
                FechaCreacion = dr.IsDBNull(13) ? DateTime.MinValue : dr.GetDateTime(13),
                FechaModificacion = dr.IsDBNull(14) ? null : dr.GetDateTime(14),
                Orientacion = dr.IsDBNull(15) ? PopupPromocionPublicoViewModel.OrientacionVertical : dr.GetString(15)
            });
        }

        return list;
    }

    public async Task<int> PopupPromocionesAdminGuardarAsync(PopupPromocionAdminFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_PopupPromociones_Guardar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@IdPopupPromocion", model.IdPopupPromocion, SqlDbType.Int);
        AddParam(cmd, "@Titulo", model.Titulo, SqlDbType.NVarChar);
        AddParam(cmd, "@Subtitulo", model.Subtitulo, SqlDbType.NVarChar);
        AddParam(cmd, "@Descripcion", model.Descripcion, SqlDbType.NVarChar);
        AddParam(cmd, "@ImagenUrl", model.ImagenUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@TextoBoton", model.TextoBoton, SqlDbType.NVarChar);
        AddParam(cmd, "@UrlBoton", model.UrlBoton, SqlDbType.NVarChar);
        AddParam(cmd, "@UrlImagen", model.UrlImagen, SqlDbType.NVarChar);
        AddParam(cmd, "@Orden", model.Orden, SqlDbType.Int);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@FechaInicio", model.FechaInicio?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaFin", model.FechaFin?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@AbrirNuevaPestana", model.AbrirNuevaPestana, SqlDbType.Bit);
        AddParam(cmd, "@Orientacion", model.Orientacion, SqlDbType.Char);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result ?? 0);
    }

    public async Task<bool> PopupPromocionesAdminEliminarAsync(int idPopupPromocion, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_PopupPromociones_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@IdPopupPromocion", idPopupPromocion, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }

    public async Task<bool> PopupPromocionesAdminCambiarEstadoAsync(int idPopupPromocion, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_PopupPromociones_CambiarEstado", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@IdPopupPromocion", idPopupPromocion, SqlDbType.Int);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }
}
