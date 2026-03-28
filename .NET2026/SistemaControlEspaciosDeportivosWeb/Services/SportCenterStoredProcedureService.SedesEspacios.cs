using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text.Json;

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
                NotificacionesActivas = ReadBool(dr, 4),
                CorreoNotificacion = dr.IsDBNull(5) ? null : dr.GetString(5),
                WhatsappContacto = dr.IsDBNull(6) ? null : dr.GetString(6),
                PermiteChatWhatsapp = ReadBool(dr, 7),
                MinutosAnticipacionRecordatorio = dr.GetInt32(8),
                MinutosToleranciaNoShow = dr.GetInt32(9),
                DiasAtencion = dr.IsDBNull(10) ? string.Empty : dr.GetString(10),
                HorarioAtencion = dr.IsDBNull(11) ? string.Empty : dr.GetString(11),
                FechasNoLaborablesCount = dr.GetInt32(12),
                Activo = ReadBool(dr, 13)
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
            Activo = ReadBool(dr, 5),
            ServiciosSeleccionados = dr.IsDBNull(6)
                ? new List<int>()
                : dr.GetString(6)
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .Distinct()
                    .ToList(),
            NotificacionesActivas = ReadBool(dr, 7),
            MinutosAnticipacionRecordatorio = dr.GetInt32(8),
            MinutosToleranciaNoShow = dr.GetInt32(9),
            CorreoNotificacion = dr.IsDBNull(10) ? null : dr.GetString(10),
            WhatsappContacto = dr.IsDBNull(11) ? null : dr.GetString(11),
            PermiteChatWhatsapp = ReadBool(dr, 12),
            AtiendeLunes = ReadBool(dr, 13),
            AtiendeMartes = ReadBool(dr, 14),
            AtiendeMiercoles = ReadBool(dr, 15),
            AtiendeJueves = ReadBool(dr, 16),
            AtiendeViernes = ReadBool(dr, 17),
            AtiendeSabado = ReadBool(dr, 18),
            AtiendeDomingo = ReadBool(dr, 19),
            HoraApertura = TimeOnly.FromTimeSpan(dr.GetTimeSpan(20)),
            HoraCierre = TimeOnly.FromTimeSpan(dr.GetTimeSpan(21)),
            FechasInhabilitadasCsv = dr.IsDBNull(22) ? null : dr.GetString(22),
            FechasInhabilitadas = dr.IsDBNull(22)
                ? new List<DateOnly>()
                : dr.GetString(22)
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Where(x => DateOnly.TryParse(x, out _))
                    .Select(DateOnly.Parse)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList()
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
        AddParam(cmd, "@AtiendeLunes", model.AtiendeLunes, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeMartes", model.AtiendeMartes, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeMiercoles", model.AtiendeMiercoles, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeJueves", model.AtiendeJueves, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeViernes", model.AtiendeViernes, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeSabado", model.AtiendeSabado, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeDomingo", model.AtiendeDomingo, SqlDbType.Bit);
        AddParam(cmd, "@HoraApertura", model.HoraApertura, SqlDbType.Time);
        AddParam(cmd, "@HoraCierre", model.HoraCierre, SqlDbType.Time);
        AddParam(cmd, "@FechasInhabilitadasCsv", model.FechasInhabilitadasCsv, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> SedesActualizarAsync(SedeFormViewModel model, string usuario)
    {
        var sedeActual = await SedesObtenerAsync(model.NegocioId, model.Id);
        if (sedeActual is null) return false;

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
        AddParam(cmd, "@AtiendeLunes", model.AtiendeLunes, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeMartes", model.AtiendeMartes, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeMiercoles", model.AtiendeMiercoles, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeJueves", model.AtiendeJueves, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeViernes", model.AtiendeViernes, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeSabado", model.AtiendeSabado, SqlDbType.Bit);
        AddParam(cmd, "@AtiendeDomingo", model.AtiendeDomingo, SqlDbType.Bit);
        AddParam(cmd, "@HoraApertura", model.HoraApertura, SqlDbType.Time);
        AddParam(cmd, "@HoraCierre", model.HoraCierre, SqlDbType.Time);
        AddParam(cmd, "@FechasInhabilitadasCsv", model.FechasInhabilitadasCsv, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();

        var sedeActualizada = await SedesObtenerAsync(model.NegocioId, model.Id);
        return sedeActualizada is not null;
    }

    public async Task<bool> SedesEliminarAsync(int negocioId, int id, string usuario)
    {
        var sedeActual = await SedesObtenerAsync(negocioId, id);
        if (sedeActual is null) return false;

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
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
                TipoSueloNombre = dr.GetString(5),
                Estado = dr.GetString(6)
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
            TipoSueloId = dr.GetInt32(3),
            Codigo = dr.GetString(4),
            Nombre = dr.GetString(5),
            Capacidad = dr.GetInt32(6),
            TieneIluminacion = ReadBool(dr, 7),
            Techada = ReadBool(dr, 8),
            Estado = (EstadoEspacioDeportivo)dr.GetInt32(9),
            Tarifas = dr.IsDBNull(10)
                ? new List<EspacioTarifaRangoViewModel>()
                : JsonSerializer.Deserialize<List<EspacioTarifaRangoViewModel>>(dr.GetString(10)) ?? new List<EspacioTarifaRangoViewModel>(),
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
        AddParam(cmd, "@TipoSueloId", model.TipoSueloId, SqlDbType.Int);
        AddParam(cmd, "@Codigo", model.Codigo, SqlDbType.NVarChar);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@Capacidad", model.Capacidad, SqlDbType.Int);
        AddParam(cmd, "@TieneIluminacion", model.TieneIluminacion, SqlDbType.Bit);
        AddParam(cmd, "@Techada", model.Techada, SqlDbType.Bit);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@TarifasJson", JsonSerializer.Serialize(model.Tarifas), SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> EspaciosActualizarAsync(EspacioFormViewModel model, string usuario)
    {
        var espacioActual = await EspaciosObtenerAsync(model.NegocioId, model.Id);
        if (espacioActual is null) return false;

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", model.SedeId, SqlDbType.Int);
        AddParam(cmd, "@TipoDeporteId", model.TipoDeporteId, SqlDbType.Int);
        AddParam(cmd, "@TipoSueloId", model.TipoSueloId, SqlDbType.Int);
        AddParam(cmd, "@Codigo", model.Codigo, SqlDbType.NVarChar);
        AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
        AddParam(cmd, "@Capacidad", model.Capacidad, SqlDbType.Int);
        AddParam(cmd, "@TieneIluminacion", model.TieneIluminacion, SqlDbType.Bit);
        AddParam(cmd, "@Techada", model.Techada, SqlDbType.Bit);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@TarifasJson", JsonSerializer.Serialize(model.Tarifas), SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();

        var espacioActualizado = await EspaciosObtenerAsync(model.NegocioId, model.Id);
        return espacioActualizado is not null;
    }

    public async Task<bool> EspaciosEliminarAsync(int negocioId, int id, string usuario)
    {
        var espacioActual = await EspaciosObtenerAsync(negocioId, id);
        if (espacioActual is null) return false;

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }

    public Task<List<SelectListItem>> EspaciosComboSedesAsync(int negocioId) => ComboAsync("Sp_Combos_Sedes", ("@NegocioId", (object?)negocioId, SqlDbType.Int));
    public Task<List<SelectListItem>> SedesComboServiciosAsync() => ComboAsync("Sp_Combos_ServiciosSede");
    public Task<List<SelectListItem>> EspaciosComboTiposDeporteAsync() => ComboAsync("Sp_Combos_TiposDeporte");
    public Task<List<SelectListItem>> EspaciosComboTiposSueloAsync() => ComboAsync("Sp_Combos_TiposSuelo");

    private static string? ToCsv(IEnumerable<int>? values)
    {
        if (values is null) return null;
        var normalized = values.Where(v => v > 0).Distinct().OrderBy(v => v).ToArray();
        return normalized.Length == 0 ? null : string.Join(",", normalized);
    }
}
