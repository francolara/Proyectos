using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public Task<List<SelectListItem>> DesafiosNivelesListarAsync()
        => ComboAsync("Sp_Desafios_Niveles_Listar");

    public async Task<List<DesafioRivalItemViewModel>> DesafiosBuscarRivalesAsync(string usuarioId, string codigoUbigeo, int? idDeporte, int? idNivel)
    {
        var list = new List<DesafioRivalItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Desafios_BuscarRivales", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@CodigoUbigeo", codigoUbigeo, SqlDbType.Char);
        AddParam(cmd, "@IdDeporte", idDeporte, SqlDbType.Int);
        AddParam(cmd, "@IdNivel", idNivel, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new DesafioRivalItemViewModel
            {
                PerfilId = dr.GetInt32(0),
                UsuarioId = dr.GetString(1),
                NombreEquipo = dr.GetString(2),
                ContactoNombre = dr.GetString(3),
                ContactoUsuario = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
                Distrito = dr.GetString(5),
                Deporte = dr.GetString(6),
                Nivel = dr.GetString(7),
                ObservacionDesafio = dr.IsDBNull(8) ? null : dr.GetString(8),
                DetalleEquipo = dr.IsDBNull(9) ? null : dr.GetString(9),
                IdDeporte = dr.IsDBNull(10) ? null : dr.GetInt32(10),
                IdNivel = dr.IsDBNull(11) ? null : dr.GetInt32(11),
                CodigoUbigeo = dr.GetString(12),
                DisponibleParaDesafios = ReadBool(dr, 13)
            });
        }

        return list;
    }

    public async Task<List<DesafioListadoItemViewModel>> DesafiosListarAsync(string usuarioId, string tipoListado)
    {
        var list = new List<DesafioListadoItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Desafios_Listar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoListado", tipoListado, SqlDbType.NVarChar);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new DesafioListadoItemViewModel
            {
                Id = dr.GetInt32(0),
                RivalNombre = dr.GetString(1),
                ContactoNombreRival = dr.GetString(2),
                ContactoUsuarioRival = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                RolVista = dr.GetString(4),
                Deporte = dr.GetString(5),
                Nivel = dr.GetString(6),
                Distrito = dr.GetString(7),
                FechaTentativa = DateOnly.FromDateTime(dr.GetDateTime(8)),
                HoraTentativa = TimeOnly.FromTimeSpan(dr.GetTimeSpan(9)),
                CanchaSugerida = dr.IsDBNull(10) ? null : dr.GetString(10),
                Modalidad = dr.GetString(11),
                Mensaje = dr.IsDBNull(12) ? null : dr.GetString(12),
                FormaPago = dr.GetString(13),
                Estado = dr.GetString(14),
                FechaCreacion = dr.GetDateTime(15),
                FechaRespuesta = dr.IsDBNull(16) ? null : dr.GetDateTime(16),
                ObservacionDesafioRival = dr.IsDBNull(17) ? null : dr.GetString(17),
                DetalleEquipoRival = dr.IsDBNull(18) ? null : dr.GetString(18),
                TelefonoRival = dr.IsDBNull(19) ? null : dr.GetString(19),
                WhatsappRival = dr.IsDBNull(20) ? null : dr.GetString(20),
                PuedeVerContactoRival = !dr.IsDBNull(21) && ReadBool(dr, 21)
            });
        }

        return list;
    }

    public async Task<(List<DesafioListadoItemViewModel> Items, int TotalRegistros)> DesafiosHistorialListarAsync(string usuarioId, int pagina = 1, int tamanoPagina = 4)
    {
        var list = new List<DesafioListadoItemViewModel>();
        var totalRegistros = 0;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Desafios_Listar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoListado", "historial", SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", pagina <= 0 ? 1 : pagina, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoPagina <= 0 ? 4 : tamanoPagina, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            if (totalRegistros == 0 && !dr.IsDBNull(22))
            {
                totalRegistros = dr.GetInt32(22);
            }

            list.Add(new DesafioListadoItemViewModel
            {
                Id = dr.GetInt32(0),
                RivalNombre = dr.GetString(1),
                ContactoNombreRival = dr.GetString(2),
                ContactoUsuarioRival = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                RolVista = dr.GetString(4),
                Deporte = dr.GetString(5),
                Nivel = dr.GetString(6),
                Distrito = dr.GetString(7),
                FechaTentativa = DateOnly.FromDateTime(dr.GetDateTime(8)),
                HoraTentativa = TimeOnly.FromTimeSpan(dr.GetTimeSpan(9)),
                CanchaSugerida = dr.IsDBNull(10) ? null : dr.GetString(10),
                Modalidad = dr.GetString(11),
                Mensaje = dr.IsDBNull(12) ? null : dr.GetString(12),
                FormaPago = dr.GetString(13),
                Estado = dr.GetString(14),
                FechaCreacion = dr.GetDateTime(15),
                FechaRespuesta = dr.IsDBNull(16) ? null : dr.GetDateTime(16),
                ObservacionDesafioRival = dr.IsDBNull(17) ? null : dr.GetString(17),
                DetalleEquipoRival = dr.IsDBNull(18) ? null : dr.GetString(18),
                TelefonoRival = dr.IsDBNull(19) ? null : dr.GetString(19),
                WhatsappRival = dr.IsDBNull(20) ? null : dr.GetString(20),
                PuedeVerContactoRival = !dr.IsDBNull(21) && ReadBool(dr, 21)
            });
        }

        return (list, totalRegistros);
    }

    public async Task<List<DesafioMensajeItemViewModel>> DesafiosMensajesListarAsync(string usuarioId, int? desafioId = null)
    {
        var list = new List<DesafioMensajeItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Desafios_Mensajes_Listar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@IdDesafio", desafioId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new DesafioMensajeItemViewModel
            {
                IdMensaje = dr.GetInt32(0),
                IdDesafio = dr.GetInt32(1),
                UsuarioIdEmisor = dr.GetString(2),
                NombreEmisor = dr.GetString(3),
                EsMio = !dr.IsDBNull(4) && ReadBool(dr, 4),
                Mensaje = dr.GetString(5),
                FechaRegistro = dr.GetDateTime(6)
            });
        }

        return list;
    }

    public async Task<int> DesafiosCrearAsync(string usuarioId, DesafioCrearViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Desafios_Crear", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@UsuarioIdRetador", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@PerfilRetadoId", model.IdUsuarioRetadoPerfil, SqlDbType.Int);
        AddParam(cmd, "@IdDeporte", model.IdDeporte, SqlDbType.Int);
        AddParam(cmd, "@IdNivel", model.IdNivel, SqlDbType.Int);
        AddParam(cmd, "@Distrito", model.Distrito, SqlDbType.Char);
        AddParam(cmd, "@FechaTentativa", model.FechaTentativa?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraTentativa", model.HoraTentativa?.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@CanchaSugerida", model.CanchaSugerida, SqlDbType.NVarChar);
        AddParam(cmd, "@Modalidad", model.Modalidad, SqlDbType.NVarChar);
        AddParam(cmd, "@Mensaje", model.Mensaje, SqlDbType.NVarChar);
        AddParam(cmd, "@FormaPago", model.FormaPago, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        var result = await cmd.ExecuteScalarAsync();
        if (result is null || result == DBNull.Value)
            throw new InvalidOperationException("El procedimiento de desafios no devolvio un identificador.");

        return Convert.ToInt32(result);
    }

    public async Task<int> DesafiosMensajeCrearAsync(string usuarioId, DesafioMensajeCrearViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Desafios_Mensajes_Crear", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@IdDesafio", model.IdDesafio, SqlDbType.Int);
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@Mensaje", model.Mensaje, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        var result = await cmd.ExecuteScalarAsync();
        if (result is null || result == DBNull.Value)
            throw new InvalidOperationException("No se pudo registrar el mensaje del desafio.");

        return Convert.ToInt32(result);
    }

    public Task<bool> DesafiosCancelarAsync(int desafioId, string usuarioId, string usuario)
        => DesafiosCambiarEstadoAsync(desafioId, usuarioId, DesafioEstados.Cancelado, usuario);

    public Task<bool> DesafiosResponderAsync(int desafioId, string usuarioId, string estado, string usuario)
        => DesafiosCambiarEstadoAsync(desafioId, usuarioId, estado, usuario);

    public Task<bool> DesafiosFinalizarAsync(int desafioId, string usuarioId, string usuario)
        => DesafiosCambiarEstadoAsync(desafioId, usuarioId, DesafioEstados.Finalizado, usuario);

    private async Task<bool> DesafiosCambiarEstadoAsync(int desafioId, string usuarioId, string estado, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Desafios_CambiarEstado", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        AddParam(cmd, "@Id", desafioId, SqlDbType.Int);
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@Estado", estado, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        await cmd.ExecuteNonQueryAsync();
        return true;
    }
}
