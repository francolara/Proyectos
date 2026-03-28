using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService(IConfiguration configuration) : ISportCenterStoredProcedureService
{
    private readonly string _connectionString = configuration.GetConnectionString("DefaultConnection")
                                               ?? throw new InvalidOperationException("No se encontro DefaultConnection.");

    private SqlConnection CreateConnection() => new(_connectionString);

    private static void AddParam(SqlCommand cmd, string name, object? value, SqlDbType type)
    {
        var p = cmd.Parameters.Add(name, type);
        p.Value = value ?? DBNull.Value;
    }

    private static bool ReadBool(SqlDataReader dr, int ordinal)
    {
        if (dr.IsDBNull(ordinal)) return false;

        var value = dr.GetValue(ordinal);
        return value switch
        {
            bool b => b,
            byte bt => bt != 0,
            short s => s != 0,
            int i => i != 0,
            long l => l != 0,
            string text when bool.TryParse(text, out var parsedBool) => parsedBool,
            string text when int.TryParse(text, out var parsedInt) => parsedInt != 0,
            _ => Convert.ToInt32(value) != 0
        };
    }

    private async Task<List<SelectListItem>> ComboAsync(string spName, params (string Name, object? Value, SqlDbType Type)[] parameters)
    {
        var list = new List<SelectListItem>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(spName, cn) { CommandType = CommandType.StoredProcedure };
        foreach (var p in parameters) AddParam(cmd, p.Name, p.Value, p.Type);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync()) list.Add(new SelectListItem(dr.GetString(1), dr.GetInt32(0).ToString()));
        return list;
    }

    public async Task<List<SedePublicaViewModel>> HomeListarSedesAsync()
    {
        var list = new List<SedePublicaViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_ListarSedesPublicas", cn) { CommandType = CommandType.StoredProcedure };
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SedePublicaViewModel
            {
                Id = dr.GetInt32(0),
                Nombre = dr.GetString(1),
                Direccion = dr.GetString(2),
                Telefono = dr.IsDBNull(3) ? null : dr.GetString(3),
                WhatsappContacto = dr.IsDBNull(4) ? null : dr.GetString(4),
                PermiteChatWhatsapp = ReadBool(dr, 5)
            });
        }
        return list;
    }

    public async Task<List<TipoDeportePublicoViewModel>> HomeListarTiposDeporteAsync()
    {
        var list = new List<TipoDeportePublicoViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_ListarTiposDeporte", cn) { CommandType = CommandType.StoredProcedure };
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new TipoDeportePublicoViewModel { Id = dr.GetInt32(0), Nombre = dr.GetString(1) });
        }
        return list;
    }

    public async Task<List<EspacioDisponibleViewModel>> HomeBuscarEspaciosDisponiblesAsync(DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin, int? sedeId, int? tipoDeporteId)
    {
        var list = new List<EspacioDisponibleViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_BuscarEspaciosDisponibles", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", horaInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", horaFin.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@TipoDeporteId", tipoDeporteId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new EspacioDisponibleViewModel
            {
                EspacioDeportivoId = dr.GetInt32(0),
                NombreEspacio = dr.GetString(1),
                Codigo = dr.GetString(2),
                SedeNombre = dr.GetString(3),
                TipoDeporteNombre = dr.GetString(4),
                TieneIluminacion = ReadBool(dr, 5),
                Techada = ReadBool(dr, 6),
                WhatsappContacto = dr.IsDBNull(7) ? null : dr.GetString(7),
                PermiteChatWhatsapp = ReadBool(dr, 8)
            });
        }
        return list;
    }

    public async Task<string> HomeSolicitarReservaPublicaAsync(SolicitudReservaPublicaFormViewModel model)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_SolicitarReservaPublica", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", model.Fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", model.HoraInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", model.HoraFin.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@NombreSolicitante", model.NombreSolicitante, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@Comentario", model.Comentario, SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return result?.ToString() ?? string.Empty;
    }

    public async Task<SolicitudPublicaDetalleViewModel?> HomeConsultarSolicitudAsync(string codigoSolicitud, string telefono)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_ConsultarSolicitudPublica", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@CodigoSolicitud", codigoSolicitud, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", telefono, SqlDbType.NVarChar);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new SolicitudPublicaDetalleViewModel
        {
            CodigoSolicitud = dr.GetString(0),
            Sede = dr.GetString(1),
            Espacio = dr.GetString(2),
            Fecha = DateOnly.FromDateTime(dr.GetDateTime(3)),
            HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(4)),
            HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(5)),
            NombreSolicitante = dr.GetString(6),
            Telefono = dr.GetString(7),
            Correo = dr.IsDBNull(8) ? null : dr.GetString(8),
            Estado = dr.GetInt32(9),
            EstadoTexto = dr.GetString(10),
            ReservaId = dr.IsDBNull(11) ? null : dr.GetInt32(11),
            FechaRegistro = dr.GetDateTime(12)
        };
    }

    public async Task<SolicitudNotificacionEmailViewModel?> HomeObtenerSolicitudParaNotificacionAsync(string codigoSolicitud)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_ObtenerSolicitudParaNotificacion", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@CodigoSolicitud", codigoSolicitud, SqlDbType.NVarChar);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new SolicitudNotificacionEmailViewModel
        {
            CodigoSolicitud = dr.GetString(0),
            NombreSolicitante = dr.GetString(1),
            Correo = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
            Sede = dr.GetString(3),
            Espacio = dr.GetString(4),
            Fecha = DateOnly.FromDateTime(dr.GetDateTime(5)),
            HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(6)),
            HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(7)),
            NotificadoCliente = ReadBool(dr, 8)
        };
    }

    public async Task<bool> HomeMarcarSolicitudNotificadaAsync(string codigoSolicitud)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_MarcarSolicitudNotificada", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@CodigoSolicitud", codigoSolicitud, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }

    public async Task<List<NegocioAccesoViewModel>> PanelListarNegociosUsuarioAsync(string usuarioId)
    {
        var list = new List<NegocioAccesoViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Panel_ListarNegociosUsuario", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new NegocioAccesoViewModel { NegocioId = dr.GetInt32(0), NombreNegocio = dr.GetString(1), Rol = dr.GetString(2) });
        }
        return list;
    }

    public async Task<string?> PanelObtenerRolAsync(string usuarioId, int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Panel_ObtenerRolUsuario", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        return (await cmd.ExecuteScalarAsync())?.ToString();
    }

    public async Task<List<PermisoModuloViewModel>> PanelListarModulosPermitidosAsync(string usuarioId, int negocioId)
    {
        var list = new List<PermisoModuloViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Panel_ListarModulosPermitidos", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@UsuarioId", usuarioId, SqlDbType.NVarChar);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new PermisoModuloViewModel
            {
                ModuloSistemaId = dr.GetInt32(0),
                ModuloCodigo = dr.GetString(1),
                ModuloNombre = dr.GetString(2),
                PuedeVer = dr.GetBoolean(3),
                PuedeCrear = dr.GetBoolean(4),
                PuedeEditar = dr.GetBoolean(5),
                PuedeEliminar = dr.GetBoolean(6)
            });
        }
        return list;
    }

    public async Task<(int TotalSedes, int TotalEspacios, int ReservasHoy, decimal IngresosHoy, decimal OcupacionHoyPct, int NoShowMes, decimal TicketPromedioMes)> PanelObtenerMetricasAsync(int negocioId, DateOnly fecha)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Panel_ObtenerMetricas", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (await dr.ReadAsync())
        {
            return (
                dr.GetInt32(0),
                dr.GetInt32(1),
                dr.GetInt32(2),
                dr.GetDecimal(3),
                dr.GetDecimal(4),
                dr.GetInt32(5),
                dr.GetDecimal(6)
            );
        }

        return (0, 0, 0, 0m, 0m, 0, 0m);
    }
}
