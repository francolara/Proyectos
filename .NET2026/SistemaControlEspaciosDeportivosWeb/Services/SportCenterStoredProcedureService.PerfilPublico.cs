using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<UsuarioPublicoPerfilViewModel?> UsuariosPublicosObtenerPerfilAsync(string usuarioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_ObtenerPerfil", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new UsuarioPublicoPerfilViewModel
        {
            Id = dr.GetInt32(0),
            UsuarioId = dr.GetString(1),
            TipoDocumento = dr.IsDBNull(2) ? "0" : dr.GetString(2),
            NumeroDocumento = dr.IsDBNull(3) ? null : dr.GetString(3),
            Nombres = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
            Apellidos = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
            NombreEquipo = dr.IsDBNull(6) ? null : dr.GetString(6),
            Telefono = dr.IsDBNull(7) ? null : dr.GetString(7),
            Correo = dr.IsDBNull(8) ? null : dr.GetString(8),
            FechaNacimiento = dr.IsDBNull(9) ? null : DateOnly.FromDateTime(dr.GetDateTime(9)),
            CodigoUbigeo = dr.IsDBNull(10) ? null : dr.GetString(10),
            CodigoDepartamento = dr.IsDBNull(11) ? null : dr.GetString(11),
            CodigoProvincia = dr.IsDBNull(12) ? null : dr.GetString(12),
            BuscarDesafios = !dr.IsDBNull(13) && ReadBool(dr, 13),
            IdDeporteDesafio = dr.IsDBNull(14) ? null : dr.GetInt32(14),
            IdNivelDesafio = dr.IsDBNull(15) ? null : dr.GetInt32(15),
            ObservacionDesafio = dr.IsDBNull(16) ? null : dr.GetString(16),
            DetalleEquipo = dr.IsDBNull(17) ? null : dr.GetString(17),
            CodigoUbigeoEquipo = dr.IsDBNull(18) ? null : dr.GetString(18),
            CodigoDepartamentoEquipo = dr.IsDBNull(19) ? null : dr.GetString(19),
            CodigoProvinciaEquipo = dr.IsDBNull(20) ? null : dr.GetString(20),
            WhatsappEquipo = dr.IsDBNull(21) ? null : dr.GetString(21)
        };
    }

    public async Task<int> UsuariosPublicosGuardarPerfilAsync(UsuarioPublicoPerfilViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_GuardarPerfil", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", model.UsuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoDocumento", model.TipoDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@NumeroDocumento", model.NumeroDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@Nombres", model.Nombres, SqlDbType.NVarChar);
        AddParam(cmd, "@Apellidos", model.Apellidos, SqlDbType.NVarChar);
        AddParam(cmd, "@NombreEquipo", model.NombreEquipo, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@FechaNacimiento", model.FechaNacimiento?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
        AddParam(cmd, "@BuscarDesafios", model.BuscarDesafios, SqlDbType.Bit);
        AddParam(cmd, "@IdDeporteDesafio", model.IdDeporteDesafio, SqlDbType.Int);
        AddParam(cmd, "@IdNivelDesafio", model.IdNivelDesafio, SqlDbType.Int);
        AddParam(cmd, "@ObservacionDesafio", model.ObservacionDesafio, SqlDbType.NVarChar);
        AddParam(cmd, "@DetalleEquipo", model.DetalleEquipo, SqlDbType.NVarChar);
        AddParam(cmd, "@CodigoUbigeoEquipo", model.CodigoUbigeoEquipo, SqlDbType.Char);
        AddParam(cmd, "@WhatsappEquipo", model.WhatsappEquipo, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result);
    }

    public async Task<(List<UsuarioPublicoReservaItemViewModel> Reservas, int TotalRegistros)> UsuariosPublicosReservasListarAsync(string usuarioId, int pagina = 1, int tamanoPagina = 6)
    {
        var list = new List<UsuarioPublicoReservaItemViewModel>();
        var totalRegistros = 0;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_ReservasListar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", pagina <= 0 ? 1 : pagina, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoPagina <= 0 ? 6 : tamanoPagina, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            if (totalRegistros == 0 && !dr.IsDBNull(18))
            {
                totalRegistros = dr.GetInt32(18);
            }

            var item = new UsuarioPublicoReservaItemViewModel
            {
                ReservaId = dr.GetInt32(0),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(1)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(2)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(3)),
                EstadoTexto = dr.GetString(4),
                Total = dr.GetDecimal(5),
                Adelanto = dr.GetDecimal(6),
                SaldoPendiente = dr.GetDecimal(7),
                NegocioNombre = dr.GetString(8),
                SedeNombre = dr.GetString(9),
                EspacioNombre = dr.GetString(10),
                SedeDireccion = dr.IsDBNull(11) ? null : dr.GetString(11),
                SedeTelefono = dr.IsDBNull(12) ? null : dr.GetString(12),
                SedeWhatsapp = dr.IsDBNull(13) ? null : dr.GetString(13),
                SedeFacebookUrl = dr.IsDBNull(14) ? null : dr.GetString(14),
                SedeInstagramUrl = dr.IsDBNull(15) ? null : dr.GetString(15),
                SedeTwitterUrl = dr.IsDBNull(16) ? null : dr.GetString(16),
                SedeMapaUrl = dr.IsDBNull(17) ? null : dr.GetString(17),
                PuedeRegistrarResena = dr.FieldCount > 19 && !dr.IsDBNull(19) && ReadBool(dr, 19)
            };

            if (dr.FieldCount > 23 && !dr.IsDBNull(20))
            {
                item.Resena = new UsuarioPublicoResenaItemViewModel
                {
                    ResenaId = dr.GetInt32(20),
                    ReservaId = item.ReservaId,
                    AliasPublico = dr.IsDBNull(21) ? string.Empty : dr.GetString(21),
                    Comentario = dr.IsDBNull(22) ? string.Empty : dr.GetString(22),
                    FechaCreacion = dr.IsDBNull(23) ? DateTime.MinValue : dr.GetDateTime(23)
                };
            }

            list.Add(item);
        }

        return (list, totalRegistros);
    }

    public async Task<UsuarioPublicoReservaCalendarioViewModel?> UsuariosPublicosReservaCalendarioObtenerAsync(string usuarioId, int reservaId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_ReservaCalendarioObtener", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@ReservaId", reservaId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new UsuarioPublicoReservaCalendarioViewModel
        {
            ReservaId = dr.GetInt32(0),
            EstadoId = dr.GetInt32(1),
            EstadoTexto = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
            NegocioNombre = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
            SedeNombre = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
            EspacioNombre = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
            SedeDireccion = dr.IsDBNull(6) ? null : dr.GetString(6),
            Fecha = DateOnly.FromDateTime(dr.GetDateTime(7)),
            HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(8)),
            HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(9))
        };
    }

    public async Task<int> UsuariosPublicosResenaCrearAsync(string usuarioId, UsuarioPublicoResenaGuardarViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_UsuariosPublicos_ResenaCrear", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@ReservaId", model.ReservaId, SqlDbType.Int);
        AddParam(cmd, "@AliasPublico", model.AliasPublico, SqlDbType.NVarChar);
        AddParam(cmd, "@Comentario", model.Comentario, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);

        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result);
    }

    public async Task<List<UsuarioPublicoResenaItemViewModel>> HomeEspacioResenasListarAsync(int espacioDeportivoId)
    {
        var list = new List<UsuarioPublicoResenaItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_EspacioResenasListar", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new UsuarioPublicoResenaItemViewModel
            {
                ResenaId = dr.GetInt32(0),
                ReservaId = dr.GetInt32(1),
                EspacioDeportivoId = dr.GetInt32(2),
                AliasPublico = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                Comentario = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
                FechaCreacion = dr.GetDateTime(5)
            });
        }

        return list;
    }
}
