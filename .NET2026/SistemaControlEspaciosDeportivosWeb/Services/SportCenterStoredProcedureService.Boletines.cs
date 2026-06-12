using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<int> BoletinesDeportivosGuardarAsync(BoletinDeportivoGuardarViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_BoletinesDeportivos_Guardar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@IdBoletin", model.IdBoletin, SqlDbType.Int);
        AddParam(cmd, "@UsuarioId", model.UsuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@Titulo", model.Titulo, SqlDbType.NVarChar);
        AddParam(cmd, "@Descripcion", model.Descripcion, SqlDbType.NVarChar);
        AddParam(cmd, "@ImagenUrl", model.ImagenUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@FechaEvento", model.FechaEvento?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
        AddParam(cmd, "@TipoRegistro", model.TipoRegistro, SqlDbType.Char);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@EsAdministradorCarga", model.EsAdministradorCarga, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        var result = await cmd.ExecuteScalarAsync();
        return result is null || result == DBNull.Value ? 0 : Convert.ToInt32(result);
    }

    public async Task<List<BoletinDeportivoPublicoItemViewModel>> BoletinesDeportivosListarPublicoAsync(
        string? codigoDepartamento = null,
        string? codigoProvincia = null,
        string? codigoUbigeo = null,
        string? zona = null,
        int? anio = null,
        int? mes = null)
    {
        var list = new List<BoletinDeportivoPublicoItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_BoletinesDeportivos_ListarPublico", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@CodigoDepartamento", string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoProvincia", string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoUbigeo", string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim(), SqlDbType.Char);
        AddParam(cmd, "@Zona", string.IsNullOrWhiteSpace(zona) ? null : zona.Trim(), SqlDbType.NVarChar);
        AddParam(cmd, "@Anio", anio, SqlDbType.Int);
        AddParam(cmd, "@Mes", mes, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(MapBoletinPublico(dr));
        }

        return list;
    }

    public async Task<(List<BoletinDeportivoUsuarioItemViewModel> Boletines, int TotalRegistros)> BoletinesDeportivosListarPorUsuarioAsync(string usuarioId, int pagina = 1, int tamanoPagina = 5)
    {
        var list = new List<BoletinDeportivoUsuarioItemViewModel>();
        var totalRegistros = 0;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_BoletinesDeportivos_ListarPorUsuario", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", pagina, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoPagina, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            var item = new BoletinDeportivoUsuarioItemViewModel();
            MapBoletinPublicoBase(dr, item);
            item.Activo = !dr.IsDBNull(11) && ReadBool(dr, 11);
            item.FechaCreacion = dr.IsDBNull(12) ? DateTime.MinValue : dr.GetDateTime(12);
            totalRegistros = dr.IsDBNull(13) ? totalRegistros : dr.GetInt32(13);
            list.Add(item);
        }

        return (list, totalRegistros);
    }

    public async Task<BoletinDeportivoDetalleViewModel?> BoletinesDeportivosObtenerPorIdAsync(int idBoletin)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_BoletinesDeportivos_ObtenerPorId", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@IdBoletin", idBoletin, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
            return null;

        return new BoletinDeportivoDetalleViewModel
        {
            IdBoletin = dr.GetInt32(0),
            UsuarioId = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
            PerfilPublicoId = dr.IsDBNull(2) ? null : dr.GetInt32(2),
            Titulo = dr.IsDBNull(3) ? null : dr.GetString(3),
            Descripcion = dr.IsDBNull(4) ? null : dr.GetString(4),
            ImagenUrl = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
            FechaEvento = DateOnly.FromDateTime(dr.GetDateTime(6)),
            CodigoUbigeo = dr.IsDBNull(7) ? string.Empty : dr.GetString(7),
            CodigoDepartamento = dr.IsDBNull(8) ? string.Empty : dr.GetString(8),
            CodigoProvincia = dr.IsDBNull(9) ? string.Empty : dr.GetString(9),
            Departamento = dr.IsDBNull(10) ? string.Empty : dr.GetString(10),
            Provincia = dr.IsDBNull(11) ? string.Empty : dr.GetString(11),
            Distrito = dr.IsDBNull(12) ? string.Empty : dr.GetString(12),
            Zona = dr.IsDBNull(13) ? null : dr.GetString(13),
            TipoRegistro = dr.IsDBNull(14) ? "U" : dr.GetString(14),
            Activo = !dr.IsDBNull(15) && ReadBool(dr, 15),
            FechaCreacion = dr.IsDBNull(16) ? DateTime.MinValue : dr.GetDateTime(16),
            UsuarioCreacion = dr.IsDBNull(17) ? string.Empty : dr.GetString(17),
            FechaActualizacion = dr.IsDBNull(18) ? null : dr.GetDateTime(18),
            UsuarioActualizacion = dr.IsDBNull(19) ? null : dr.GetString(19)
        };
    }

    public async Task<BoletinesDeportivosAdminResumenViewModel> BoletinesDeportivosAdminResumenAsync()
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_BoletinesDeportivos_AdminResumen", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
            return new BoletinesDeportivosAdminResumenViewModel();

        return new BoletinesDeportivosAdminResumenViewModel
        {
            TotalBoletines = dr.IsDBNull(0) ? 0 : dr.GetInt32(0),
            TotalActivos = dr.IsDBNull(1) ? 0 : dr.GetInt32(1),
            TotalInactivos = dr.IsDBNull(2) ? 0 : dr.GetInt32(2),
            TotalUsuarios = dr.IsDBNull(3) ? 0 : dr.GetInt32(3),
            TotalPlataforma = dr.IsDBNull(4) ? 0 : dr.GetInt32(4)
        };
    }

    public async Task<(List<BoletinDeportivoAdminItemViewModel> Boletines, int TotalRegistros)> BoletinesDeportivosAdminListarAsync(
        bool? activo = null,
        string? tipoRegistro = null,
        string? codigoDepartamento = null,
        string? codigoProvincia = null,
        string? codigoUbigeo = null,
        string? zona = null,
        int? anio = null,
        int? mes = null,
        int pagina = 1,
        int tamanoPagina = 5)
    {
        var list = new List<BoletinDeportivoAdminItemViewModel>();
        var totalRegistros = 0;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_BoletinesDeportivos_AdminListar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@TipoRegistro", string.IsNullOrWhiteSpace(tipoRegistro) ? null : tipoRegistro.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoDepartamento", string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoProvincia", string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoUbigeo", string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim(), SqlDbType.Char);
        AddParam(cmd, "@Zona", string.IsNullOrWhiteSpace(zona) ? null : zona.Trim(), SqlDbType.NVarChar);
        AddParam(cmd, "@Anio", anio, SqlDbType.Int);
        AddParam(cmd, "@Mes", mes, SqlDbType.Int);
        AddParam(cmd, "@Pagina", pagina, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoPagina, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            var item = new BoletinDeportivoAdminItemViewModel
            {
                IdBoletin = dr.GetInt32(0),
                UsuarioId = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                CorreoAutor = dr.IsDBNull(2) ? null : dr.GetString(2),
                NombreAutor = dr.IsDBNull(3) ? string.Empty : dr.GetString(3)
            };
            item.Titulo = dr.IsDBNull(4) ? null : dr.GetString(4);
            item.Descripcion = dr.IsDBNull(5) ? null : dr.GetString(5);
            item.ImagenUrl = dr.IsDBNull(6) ? string.Empty : dr.GetString(6);
            item.FechaEvento = DateOnly.FromDateTime(dr.GetDateTime(7));
            item.CodigoUbigeo = dr.IsDBNull(8) ? string.Empty : dr.GetString(8);
            item.Departamento = dr.IsDBNull(9) ? string.Empty : dr.GetString(9);
            item.Provincia = dr.IsDBNull(10) ? string.Empty : dr.GetString(10);
            item.Distrito = dr.IsDBNull(11) ? string.Empty : dr.GetString(11);
            item.Zona = dr.IsDBNull(12) ? null : dr.GetString(12);
            item.TipoRegistro = dr.IsDBNull(13) ? "U" : dr.GetString(13);
            item.Activo = !dr.IsDBNull(14) && ReadBool(dr, 14);
            item.FechaCreacion = dr.IsDBNull(15) ? DateTime.MinValue : dr.GetDateTime(15);
            totalRegistros = dr.IsDBNull(16) ? totalRegistros : dr.GetInt32(16);
            list.Add(item);
        }

        return (list, totalRegistros);
    }

    public async Task<bool> BoletinesDeportivosCambiarEstadoAsync(int idBoletin, bool activo, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_BoletinesDeportivos_CambiarEstado", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@IdBoletin", idBoletin, SqlDbType.Int);
        AddParam(cmd, "@Activo", activo, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return result is not null && result != DBNull.Value && Convert.ToBoolean(result);
    }

    private static BoletinDeportivoPublicoItemViewModel MapBoletinPublico(SqlDataReader dr)
    {
        var item = new BoletinDeportivoPublicoItemViewModel();
        MapBoletinPublicoBase(dr, item);
        item.FechaCreacion = dr.IsDBNull(11) ? DateTime.MinValue : dr.GetDateTime(11);
        return item;
    }

    private static void MapBoletinPublicoBase(SqlDataReader dr, BoletinDeportivoPublicoItemViewModel item)
    {
        item.IdBoletin = dr.GetInt32(0);
        item.Titulo = dr.IsDBNull(1) ? null : dr.GetString(1);
        item.Descripcion = dr.IsDBNull(2) ? null : dr.GetString(2);
        item.ImagenUrl = dr.IsDBNull(3) ? string.Empty : dr.GetString(3);
        item.FechaEvento = DateOnly.FromDateTime(dr.GetDateTime(4));
        item.CodigoUbigeo = dr.IsDBNull(5) ? string.Empty : dr.GetString(5);
        item.Departamento = dr.IsDBNull(6) ? string.Empty : dr.GetString(6);
        item.Provincia = dr.IsDBNull(7) ? string.Empty : dr.GetString(7);
        item.Distrito = dr.IsDBNull(8) ? string.Empty : dr.GetString(8);
        item.Zona = dr.IsDBNull(9) ? null : dr.GetString(9);
        item.TipoRegistro = dr.FieldCount > 10 && !dr.IsDBNull(10) ? dr.GetString(10) : "U";
    }
}
