using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text.Json;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    private static readonly JsonSerializerOptions TarifaJsonSerializerOptions = new(JsonSerializerDefaults.Web);

    public async Task<List<SedeItemViewModel>> SedesListarAsync(int negocioId, int? sedeId = null)
    {
        var list = new List<SedeItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
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
            CodigoUbigeo = dr.FieldCount > 33 && !dr.IsDBNull(33) ? dr.GetString(33) : string.Empty,
            CodigoDepartamento = dr.FieldCount > 34 && !dr.IsDBNull(34) ? dr.GetString(34) : null,
            CodigoProvincia = dr.FieldCount > 35 && !dr.IsDBNull(35) ? dr.GetString(35) : null,
            ConsideracionesReserva = dr.FieldCount > 4 && !dr.IsDBNull(4) ? dr.GetString(4) : null,
            Telefono = dr.IsDBNull(5) ? null : dr.GetString(5),
            FacebookUrl = dr.FieldCount > 6 && !dr.IsDBNull(6) ? dr.GetString(6) : null,
            InstagramUrl = dr.FieldCount > 7 && !dr.IsDBNull(7) ? dr.GetString(7) : null,
            TwitterUrl = dr.FieldCount > 8 && !dr.IsDBNull(8) ? dr.GetString(8) : null,
            Activo = ReadBool(dr, 9),
            Latitud = dr.FieldCount > 10 && !dr.IsDBNull(10) ? dr.GetDecimal(10) : null,
            Longitud = dr.FieldCount > 11 && !dr.IsDBNull(11) ? dr.GetDecimal(11) : null,
            GooglePlaceId = dr.FieldCount > 12 && !dr.IsDBNull(12) ? dr.GetString(12) : null,
            GoogleMapsUrl = dr.FieldCount > 13 && !dr.IsDBNull(13) ? dr.GetString(13) : null,
            FotoPrincipalUrl = dr.FieldCount > 14 && !dr.IsDBNull(14) ? dr.GetString(14) : null,
            FotosUrlsCsv = dr.FieldCount > 15 && !dr.IsDBNull(15) ? dr.GetString(15) : null,
            FotosUrls = dr.FieldCount > 15 && !dr.IsDBNull(15)
                ? dr.GetString(15)
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList()
                : new List<string>(),
            ServiciosSeleccionados = dr.FieldCount > 16 && !dr.IsDBNull(16)
                ? dr.GetString(16)
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .Distinct()
                    .ToList()
                : new List<int>(),
            NotificacionesActivas = dr.FieldCount > 17 ? ReadBool(dr, 17) : true,
            MinutosAnticipacionRecordatorio = dr.FieldCount > 18 && !dr.IsDBNull(18) ? dr.GetInt32(18) : 90,
            MinutosToleranciaNoShow = dr.FieldCount > 19 && !dr.IsDBNull(19) ? dr.GetInt32(19) : 30,
            CorreoNotificacion = dr.FieldCount > 20 && !dr.IsDBNull(20) ? dr.GetString(20) : null,
            WhatsappContacto = dr.FieldCount > 21 && !dr.IsDBNull(21) ? dr.GetString(21) : null,
            PermiteChatWhatsapp = dr.FieldCount > 22 && ReadBool(dr, 22),
            AtiendeLunes = dr.FieldCount > 23 ? ReadBool(dr, 23) : true,
            AtiendeMartes = dr.FieldCount > 24 ? ReadBool(dr, 24) : true,
            AtiendeMiercoles = dr.FieldCount > 25 ? ReadBool(dr, 25) : true,
            AtiendeJueves = dr.FieldCount > 26 ? ReadBool(dr, 26) : true,
            AtiendeViernes = dr.FieldCount > 27 ? ReadBool(dr, 27) : true,
            AtiendeSabado = dr.FieldCount > 28 ? ReadBool(dr, 28) : true,
            AtiendeDomingo = dr.FieldCount > 29 ? ReadBool(dr, 29) : true,
            HoraApertura = dr.FieldCount > 30 && !dr.IsDBNull(30) ? TimeOnly.FromTimeSpan(dr.GetTimeSpan(30)) : new TimeOnly(8, 0),
            HoraCierre = dr.FieldCount > 31 && !dr.IsDBNull(31) ? TimeOnly.FromTimeSpan(dr.GetTimeSpan(31)) : new TimeOnly(23, 0),
            FechasInhabilitadasCsv = dr.FieldCount > 32 && !dr.IsDBNull(32) ? dr.GetString(32) : null,
            FechasInhabilitadas = dr.FieldCount > 32 && !dr.IsDBNull(32)
                ? dr.GetString(32)
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Where(x => DateOnly.TryParse(x, out _))
                    .Select(DateOnly.Parse)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList()
                : new List<DateOnly>()
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
        AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
        AddParam(cmd, "@ConsideracionesReserva", model.ConsideracionesReserva, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@FacebookUrl", model.FacebookUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@InstagramUrl", model.InstagramUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@TwitterUrl", model.TwitterUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
        AddParam(cmd, "@Latitud", model.Latitud, SqlDbType.Decimal);
        AddParam(cmd, "@Longitud", model.Longitud, SqlDbType.Decimal);
        AddParam(cmd, "@GooglePlaceId", model.GooglePlaceId, SqlDbType.NVarChar);
        AddParam(cmd, "@GoogleDepartamento", model.GoogleDepartamento, SqlDbType.NVarChar);
        AddParam(cmd, "@GoogleProvincia", model.GoogleProvincia, SqlDbType.NVarChar);
        AddParam(cmd, "@GoogleDistrito", model.GoogleDistrito, SqlDbType.NVarChar);
        AddParam(cmd, "@GoogleMapsUrl", model.GoogleMapsUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@FotoPrincipalUrl", model.FotoPrincipalUrl, SqlDbType.NVarChar);
        AddParam(cmd, "@FotosUrlsCsv", model.FotosUrlsCsv, SqlDbType.NVarChar);
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
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Sedes_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@Nombre", model.Nombre, SqlDbType.NVarChar);
            AddParam(cmd, "@Direccion", model.Direccion, SqlDbType.NVarChar);
            AddParam(cmd, "@CodigoUbigeo", model.CodigoUbigeo, SqlDbType.Char);
            AddParam(cmd, "@ConsideracionesReserva", model.ConsideracionesReserva, SqlDbType.NVarChar);
            AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
            AddParam(cmd, "@FacebookUrl", model.FacebookUrl, SqlDbType.NVarChar);
            AddParam(cmd, "@InstagramUrl", model.InstagramUrl, SqlDbType.NVarChar);
            AddParam(cmd, "@TwitterUrl", model.TwitterUrl, SqlDbType.NVarChar);
            AddParam(cmd, "@Activo", model.Activo, SqlDbType.Bit);
            AddParam(cmd, "@Latitud", model.Latitud, SqlDbType.Decimal);
            AddParam(cmd, "@Longitud", model.Longitud, SqlDbType.Decimal);
            AddParam(cmd, "@GooglePlaceId", model.GooglePlaceId, SqlDbType.NVarChar);
            AddParam(cmd, "@GoogleDepartamento", model.GoogleDepartamento, SqlDbType.NVarChar);
            AddParam(cmd, "@GoogleProvincia", model.GoogleProvincia, SqlDbType.NVarChar);
            AddParam(cmd, "@GoogleDistrito", model.GoogleDistrito, SqlDbType.NVarChar);
            AddParam(cmd, "@GoogleMapsUrl", model.GoogleMapsUrl, SqlDbType.NVarChar);
            AddParam(cmd, "@FotoPrincipalUrl", model.FotoPrincipalUrl, SqlDbType.NVarChar);
            AddParam(cmd, "@FotosUrlsCsv", model.FotosUrlsCsv, SqlDbType.NVarChar);
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
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> SedesEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Sedes_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<List<EspacioItemViewModel>> EspaciosListarAsync(int negocioId, int? sedeId = null)
    {
        var list = new List<EspacioItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Espacios_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (dr.FieldCount < 10)
            throw new InvalidOperationException("El SP Sp_Espacios_Listar no devuelve la columna TarifaResumen. Ejecuta el script SQL actualizado.");
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
                TieneIluminacion = ReadBool(dr, 6),
                Techada = ReadBool(dr, 7),
                Estado = dr.GetString(8),
                TarifaResumen = dr.IsDBNull(9)
                    ? "Sin tarifa configurada"
                    : dr.GetString(9),
                AdministracionPrivada = dr.FieldCount > 10 && ReadBool(dr, 10)
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
                : JsonSerializer.Deserialize<List<EspacioTarifaRangoViewModel>>(dr.GetString(10), TarifaJsonSerializerOptions) ?? new List<EspacioTarifaRangoViewModel>(),
            AdministracionPrivada = dr.FieldCount > 11 && ReadBool(dr, 11),
            TarifasFeriado = dr.FieldCount > 12 && !dr.IsDBNull(12)
                ? JsonSerializer.Deserialize<List<EspacioTarifaFeriadoRangoViewModel>>(dr.GetString(12), TarifaJsonSerializerOptions) ?? new List<EspacioTarifaFeriadoRangoViewModel>()
                : new List<EspacioTarifaFeriadoRangoViewModel>(),
            NegocioId = negocioId
        };
    }

    public async Task<int> EspaciosCrearAsync(EspacioFormViewModel model, string usuario)
    {
        try
        {
            return await EspaciosCrearInternoAsync(model, usuario, incluirTarifasFeriado: true);
        }
        catch (SqlException ex) when (ex.Message.Contains("@TarifasFeriadoJson", StringComparison.OrdinalIgnoreCase))
        {
            return await EspaciosCrearInternoAsync(model, usuario, incluirTarifasFeriado: false);
        }
    }

    public async Task<bool> EspaciosActualizarAsync(EspacioFormViewModel model, string usuario)
    {
        try
        {
            await EspaciosActualizarInternoAsync(model, usuario, incluirTarifasFeriado: true);
            return true;
        }
        catch (SqlException ex) when (ex.Message.Contains("@TarifasFeriadoJson", StringComparison.OrdinalIgnoreCase))
        {
            await EspaciosActualizarInternoAsync(model, usuario, incluirTarifasFeriado: false);
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    private async Task<int> EspaciosCrearInternoAsync(EspacioFormViewModel model, string usuario, bool incluirTarifasFeriado)
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
        AddParam(cmd, "@AdministracionPrivada", model.AdministracionPrivada, SqlDbType.Bit);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@TarifasJson", ObtenerTarifasJson(model), SqlDbType.NVarChar);
        if (incluirTarifasFeriado)
            AddParam(cmd, "@TarifasFeriadoJson", ObtenerTarifasFeriadoJson(model), SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    private async Task EspaciosActualizarInternoAsync(EspacioFormViewModel model, string usuario, bool incluirTarifasFeriado)
    {
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
        AddParam(cmd, "@AdministracionPrivada", model.AdministracionPrivada, SqlDbType.Bit);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@TarifasJson", ObtenerTarifasJson(model), SqlDbType.NVarChar);
        if (incluirTarifasFeriado)
            AddParam(cmd, "@TarifasFeriadoJson", ObtenerTarifasFeriadoJson(model), SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
    }

    public async Task<bool> EspaciosEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Espacios_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public Task<List<SelectListItem>> EspaciosComboSedesAsync(int negocioId, int? sedeId = null) => ComboAsync("Sp_Combos_Sedes", ("@NegocioId", (object?)negocioId, SqlDbType.Int), ("@SedeId", sedeId, SqlDbType.Int));
    public Task<List<SelectListItem>> SedesComboServiciosAsync() => ComboAsync("Sp_Combos_ServiciosSede");
    public Task<List<SelectListItem>> EspaciosComboTiposDeporteAsync(int negocioId) => ComboAsync("Sp_Combos_TiposDeporte", ("@NegocioId", negocioId, SqlDbType.Int));
    public Task<List<SelectListItem>> EspaciosComboTiposSueloAsync(int negocioId) => ComboAsync("Sp_Combos_TiposSuelo", ("@NegocioId", negocioId, SqlDbType.Int));

    public async Task<List<SedeSerieDocumentoConfigItemViewModel>> SedesSeriesDocumentoListarAsync(int negocioId, int sedeId)
    {
        var list = new List<SedeSerieDocumentoConfigItemViewModel>();
        var lookup = new Dictionary<string, SedeSerieDocumentoConfigItemViewModel>(StringComparer.OrdinalIgnoreCase);
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_SeriesDocumentoComprobante_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            var codigo = dr.GetString(0);
            if (!lookup.TryGetValue(codigo, out var item))
            {
                item = new SedeSerieDocumentoConfigItemViewModel
                {
                    CodigoSunat = codigo,
                    NombreDocumento = dr.GetString(1),
                    Tributario = dr.GetBoolean(2),
                    SerieSeleccionada = dr.IsDBNull(4) ? null : dr.GetString(4)
                };
                lookup[codigo] = item;
                list.Add(item);
            }

            if (!dr.IsDBNull(3))
            {
                var negocioSerieId = dr.GetInt32(3);
                if (!item.NegocioSeriesIds.Contains(negocioSerieId))
                    item.NegocioSeriesIds.Add(negocioSerieId);
            }
        }

        foreach (var item in list)
        {
            item.SeriesDisponibles = await CombosSeriesDocumentoComprobanteAsync(negocioId, item.CodigoSunat);
            if (!item.PermiteMultiplesSeries)
                item.NegocioSerieId = item.NegocioSeriesIds.FirstOrDefault() > 0 ? item.NegocioSeriesIds.First() : null;
        }

        return list;
    }

    public async Task SedesSeriesDocumentoGuardarAsync(int negocioId, int sedeId, string codigoSunat, int? negocioSerieId, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_SeriesDocumentoComprobante_Guardar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@CodigoSunat", codigoSunat, SqlDbType.NVarChar);
        AddParam(cmd, "@NegocioSerieId", negocioSerieId, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
    }

    public async Task SedesSeriesDocumentoGuardarMultiplesAsync(int negocioId, int sedeId, string codigoSunat, IEnumerable<int> negocioSeriesIds, string usuario)
    {
        var csvSeries = ToCsv(negocioSeriesIds);
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Sedes_SeriesDocumentoComprobante_Guardar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@CodigoSunat", codigoSunat, SqlDbType.NVarChar);
        AddParam(cmd, "@NegocioSerieId", null, SqlDbType.Int);
        AddParam(cmd, "@NegocioSeriesIdsCsv", csvSeries, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
    }

    private static string? ToCsv(IEnumerable<int>? values)
    {
        if (values is null) return null;
        var normalized = values.Where(v => v > 0).Distinct().OrderBy(v => v).ToArray();
        return normalized.Length == 0 ? null : string.Join(",", normalized);
    }

    private static string ObtenerTarifasJson(EspacioFormViewModel model)
        => string.IsNullOrWhiteSpace(model.TarifasJson)
            ? JsonSerializer.Serialize(model.Tarifas, TarifaJsonSerializerOptions)
            : model.TarifasJson;

    private static string ObtenerTarifasFeriadoJson(EspacioFormViewModel model)
        => string.IsNullOrWhiteSpace(model.TarifasFeriadoJson)
            ? JsonSerializer.Serialize(model.TarifasFeriado, TarifaJsonSerializerOptions)
            : model.TarifasFeriadoJson;
}
