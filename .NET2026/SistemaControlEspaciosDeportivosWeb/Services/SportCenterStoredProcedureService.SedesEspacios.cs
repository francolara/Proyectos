using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<SedeItemViewModel>> SedesListarAsync(int negocioId)
    {
        var list = new List<SedeItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new SedeItemViewModel
            {
                Id = dr.GetInt32(0),
                Nombre = dr.GetString(1),
                Direccion = dr.GetString(2),
                Servicios = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                NotificacionesActivas = dr.GetBoolean(4),
                CorreoNotificacion = dr.IsDBNull(5) ? null : dr.GetString(5),
                WhatsappContacto = dr.IsDBNull(6) ? null : dr.GetString(6),
                PermiteChatWhatsapp = dr.GetBoolean(7),
                MinutosAnticipacionRecordatorio = dr.GetInt32(8),
                MinutosToleranciaNoShow = dr.GetInt32(9),
                Activo = dr.GetBoolean(10)
            });
        }
        return list;
    }

    public async Task<SedeFormViewModel?> SedesObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        return new SedeFormViewModel
        {
            Id = dr.GetInt32(0),
            NegocioId = dr.GetInt32(1),
            Nombre = dr.GetString(2),
            Direccion = dr.GetString(3),
            Telefono = dr.IsDBNull(4) ? null : dr.GetString(4),
            Activo = dr.GetBoolean(5),
            ServiciosSeleccionados = dr.IsDBNull(6)
                ? new List<int>()
                : dr.GetString(6)
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .Distinct()
                    .ToList(),
            NotificacionesActivas = dr.GetBoolean(7),
            MinutosAnticipacionRecordatorio = dr.GetInt32(8),
            MinutosToleranciaNoShow = dr.GetInt32(9),
            CorreoNotificacion = dr.IsDBNull(10) ? null : dr.GetString(10),
            WhatsappContacto = dr.IsDBNull(11) ? null : dr.GetString(11),
            PermiteChatWhatsapp = dr.GetBoolean(12)
        };
    }

    public async Task<int> SedesCrearAsync(SedeFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@Direccion", model.Direccion, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@ServiciosIdsCsv", ToCsv(model.ServiciosSeleccionados), SqlDbType.NVarChar);
        AddParam(cmd, "@NotificacionesActivas", model.NotificacionesActivas, SqlDbType.Bit);
        AddParam(cmd, "@MinutosAnticipacionRecordatorio", model.MinutosAnticipacionRecordatorio, SqlDbType.Int);
        AddParam(cmd, "@MinutosToleranciaNoShow", model.MinutosToleranciaNoShow, SqlDbType.Int);
        AddParam(cmd, "@CorreoNotificacion", model.CorreoNotificacion, SqlDbType.NVarChar);
        AddParam(cmd, "@WhatsappContacto", model.WhatsappContacto, SqlDbType.NVarChar);
        AddParam(cmd, "@PermiteChatWhatsapp", model.PermiteChatWhatsapp, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> SedesActualizarAsync(SedeFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@Direccion", model.Direccion, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@ServiciosIdsCsv", ToCsv(model.ServiciosSeleccionados), SqlDbType.NVarChar);
        AddParam(cmd, "@NotificacionesActivas", model.NotificacionesActivas, SqlDbType.Bit);
        AddParam(cmd, "@MinutosAnticipacionRecordatorio", model.MinutosAnticipacionRecordatorio, SqlDbType.Int);
        AddParam(cmd, "@MinutosToleranciaNoShow", model.MinutosToleranciaNoShow, SqlDbType.Int);
        AddParam(cmd, "@CorreoNotificacion", model.CorreoNotificacion, SqlDbType.NVarChar);
        AddParam(cmd, "@WhatsappContacto", model.WhatsappContacto, SqlDbType.NVarChar);
        AddParam(cmd, "@PermiteChatWhatsapp", model.PermiteChatWhatsapp, SqlDbType.Bit);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> SedesEliminarAsync(int negocioId, int id, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<List<EspacioItemViewModel>> EspaciosListarAsync(int negocioId)
    {
        var list = new List<EspacioItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new EspacioItemViewModel
            {
                Id = dr.GetInt32(0),
                Codigo = dr.GetString(1),
                Nombre = dr.GetString(2),
                SedeNombre = dr.GetString(3),
                TipoDeporteNombre = dr.GetString(4),
                Estado = dr.GetString(5)
            });
        }
        return list;
    }

    public async Task<EspacioFormViewModel?> EspaciosObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        return new EspacioFormViewModel
        {
            Id = dr.GetInt32(0),
            SedeId = dr.GetInt32(1),
            TipoDeporteId = dr.GetInt32(2),
            Codigo = dr.GetString(3),
            Nombre = dr.GetString(4),
            Capacidad = dr.GetInt32(5),
            TieneIluminacion = dr.GetBoolean(6),
            Techada = dr.GetBoolean(7),
            Estado = (EstadoEspacioDeportivo)dr.GetInt32(8),
            NegocioId = negocioId
        };
    }

    public async Task<int> EspaciosCrearAsync(EspacioFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", model.SedeId, SqlDbType.Int);
        AddParam(cmd, "@TipoDeporteId", model.TipoDeporteId, SqlDbType.Int);
        AddParam(cmd, "@Codigo", model.Codigo, SqlDbType.NVarChar);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@Capacidad", model.Capacidad, SqlDbType.Int);
        AddParam(cmd, "@TieneIluminacion", model.TieneIluminacion, SqlDbType.Bit);
        AddParam(cmd, "@Techada", model.Techada, SqlDbType.Bit);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> EspaciosActualizarAsync(EspacioFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", model.SedeId, SqlDbType.Int);
        AddParam(cmd, "@TipoDeporteId", model.TipoDeporteId, SqlDbType.Int);
        AddParam(cmd, "@Codigo", model.Codigo, SqlDbType.NVarChar);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@Capacidad", model.Capacidad, SqlDbType.Int);
        AddParam(cmd, "@TieneIluminacion", model.TieneIluminacion, SqlDbType.Bit);
        AddParam(cmd, "@Techada", model.Techada, SqlDbType.Bit);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> EspaciosEliminarAsync(int negocioId, int id, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public Task<List<SelectListItem>> EspaciosComboSedesAsync(int negocioId) => ComboAsync("Sp_Combos_Sedes", ("@NegocioId", (object?)negocioId, SqlDbType.Int));
    public Task<List<SelectListItem>> SedesComboServiciosAsync() => ComboAsync("Sp_Combos_ServiciosSede");
    public Task<List<SelectListItem>> EspaciosComboTiposDeporteAsync() => ComboAsync("Sp_Combos_TiposDeporte");

    private static string? ToCsv(IEnumerable<int>? values)
    {
        if (values is null) return null;
        var normalized = values.Where(v => v > 0).Distinct().OrderBy(v => v).ToArray();
        return normalized.Length == 0 ? null : string.Join(",", normalized);
    }
}
