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
        while (await dr.ReadAsync())
        {
            var text = dr.IsDBNull(1) ? string.Empty : dr.GetString(1);
            var rawValue = dr.IsDBNull(0) ? string.Empty : Convert.ToString(dr.GetValue(0)) ?? string.Empty;
            list.Add(new SelectListItem(text, rawValue));
        }
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
                ConsideracionesReserva = dr.FieldCount > 3 && !dr.IsDBNull(3) ? dr.GetString(3) : null,
                Telefono = dr.IsDBNull(4) ? null : dr.GetString(4),
                WhatsappContacto = dr.IsDBNull(5) ? null : dr.GetString(5),
                PermiteChatWhatsapp = ReadBool(dr, 6),
                FacebookUrl = dr.FieldCount > 7 && !dr.IsDBNull(7) ? dr.GetString(7) : null,
                InstagramUrl = dr.FieldCount > 8 && !dr.IsDBNull(8) ? dr.GetString(8) : null,
                TwitterUrl = dr.FieldCount > 9 && !dr.IsDBNull(9) ? dr.GetString(9) : null,
                Latitud = dr.FieldCount > 10 && !dr.IsDBNull(10) ? dr.GetDecimal(10) : null,
                Longitud = dr.FieldCount > 11 && !dr.IsDBNull(11) ? dr.GetDecimal(11) : null,
                GoogleMapsUrl = dr.FieldCount > 12 && !dr.IsDBNull(12) ? dr.GetString(12) : null,
                FotoPrincipalUrl = dr.FieldCount > 13 && !dr.IsDBNull(13) ? dr.GetString(13) : null,
                FotosAlternativas = dr.FieldCount > 14 && !dr.IsDBNull(14)
                    ? dr.GetString(14)
                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList()
                    : new List<string>(),
                NegocioId = dr.FieldCount > 15 && !dr.IsDBNull(15) ? dr.GetInt32(15) : null,
                NegocioNombre = dr.FieldCount > 16 && !dr.IsDBNull(16) ? dr.GetString(16) : null,
                Servicios = dr.FieldCount > 17 && !dr.IsDBNull(17) ? dr.GetString(17) : null,
                CodigoUbigeoNegocio = dr.FieldCount > 18 && !dr.IsDBNull(18) ? dr.GetString(18) : null,
                CodigoDepartamentoNegocio = dr.FieldCount > 19 && !dr.IsDBNull(19) ? dr.GetString(19) : null,
                CodigoProvinciaNegocio = dr.FieldCount > 20 && !dr.IsDBNull(20) ? dr.GetString(20) : null
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

    public async Task<List<SelectListItem>> HomeReferencialesExternosListarTiposDeporteSuperAsync()
    {
        return await ComboAsync("Sp_Home_ReferencialesExternos_ListarTiposDeporteSuper");
    }

    public async Task<(List<ReferencialExternoAdminItemViewModel> Items, int TotalRegistros)> HomeReferencialesExternosListarAdminAsync(
        string? codigoDepartamento = null,
        string? codigoProvincia = null,
        string? codigoUbigeo = null,
        string? buscarNombre = null,
        int pagina = 1,
        int tamanoPagina = 20,
        bool? soloActivos = true)
    {
        var list = new List<ReferencialExternoAdminItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_ReferencialesExternos_ListarAdmin", cn) { CommandType = CommandType.StoredProcedure };

        AddParam(cmd, "@CodigoDepartamento", string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoProvincia", string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoUbigeo", string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim(), SqlDbType.Char);
        AddParam(cmd, "@BuscarNombre", string.IsNullOrWhiteSpace(buscarNombre) ? null : buscarNombre.Trim(), SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", pagina <= 0 ? 1 : pagina, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoPagina <= 0 ? 20 : tamanoPagina, SqlDbType.Int);
        AddParam(cmd, "@SoloActivos", soloActivos, SqlDbType.Bit);

        var totalRegistrosParam = cmd.Parameters.Add("@TotalRegistros", SqlDbType.Int);
        totalRegistrosParam.Direction = ParameterDirection.Output;

        await using (var dr = await cmd.ExecuteReaderAsync())
        {
            while (await dr.ReadAsync())
            {
                list.Add(new ReferencialExternoAdminItemViewModel
                {
                Id = dr.GetInt32(0),
                NombreComplejo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                NombreEspacio = dr.IsDBNull(2) ? null : dr.GetString(2),
                CodigoUbigeo = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                TipoDeporteSuperId = dr.IsDBNull(4) ? 0 : dr.GetInt32(4),
                TipoDeporte = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
                Departamento = dr.IsDBNull(6) ? string.Empty : dr.GetString(6),
                Provincia = dr.IsDBNull(7) ? string.Empty : dr.GetString(7),
                Distrito = dr.IsDBNull(8) ? string.Empty : dr.GetString(8),
                Direccion = dr.IsDBNull(9) ? null : dr.GetString(9),
                TelefonoContacto = dr.IsDBNull(10) ? null : dr.GetString(10),
                GoogleMapsUrl = dr.IsDBNull(11) ? null : dr.GetString(11),
                Activo = ReadBool(dr, 12),
                FechaActualizacion = dr.IsDBNull(13) ? null : dr.GetDateTime(13),
                UsuarioActualizacion = dr.IsDBNull(14) ? null : dr.GetString(14)
            });
        }
        }

        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        return (list, totalRegistros);
    }

    public async Task<bool> HomeReferencialesExternosInactivarAsync(int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Home_ReferencialesExternos_Inactivar", cn) { CommandType = CommandType.StoredProcedure };
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

    public async Task<bool> HomeReferencialesExternosActivarAsync(int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Home_ReferencialesExternos_Activar", cn) { CommandType = CommandType.StoredProcedure };
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

    public async Task<bool> HomeReferencialesExternosActualizarAsync(int id, string nombreComplejo, string? telefonoContacto, int tipoDeporteSuperId, string? direccion, string codigoUbigeo, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Home_ReferencialesExternos_ActualizarAdmin", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@NombreComplejo", nombreComplejo.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@TelefonoContacto", string.IsNullOrWhiteSpace(telefonoContacto) ? null : telefonoContacto.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@TipoDeporteSuperId", tipoDeporteSuperId, SqlDbType.Int);
            AddParam(cmd, "@Direccion", string.IsNullOrWhiteSpace(direccion) ? null : direccion.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@CodigoUbigeo", codigoUbigeo.Trim(), SqlDbType.Char);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<List<EspacioDisponibleViewModel>> HomeBuscarEspaciosDisponiblesAsync(
        DateOnly fecha,
        TimeOnly horaInicio,
        TimeOnly horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId,
        int? negocioId,
        bool omitirFechaHorario = false,
        bool buscarCercaDeMi = false,
        decimal? latitudUsuario = null,
        decimal? longitudUsuario = null,
        decimal? radioKm = null)
    {
        try
        {
            return await HomeBuscarEspaciosDisponiblesInternoAsync(
                fecha,
                horaInicio,
                horaFin,
                codigoDepartamento,
                codigoProvincia,
                codigoUbigeo,
                tipoDeporteId,
                negocioId,
                omitirFechaHorario,
                buscarCercaDeMi,
                latitudUsuario,
                longitudUsuario,
                radioKm,
                usarUbigeo: true,
                usarNegocio: true,
                usarIgnorarHorario: true,
                usarFiltroDistancia: true);
        }
        catch (SqlException ex) when (
            ex.Message.Contains("@CodigoDepartamento", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("@CodigoProvincia", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("@CodigoUbigeo", StringComparison.OrdinalIgnoreCase))
        {
            // Compatibilidad temporal con SP antiguo (filtro por sede).
            return await HomeBuscarEspaciosDisponiblesInternoAsync(
                fecha,
                horaInicio,
                horaFin,
                null,
                null,
                null,
                tipoDeporteId,
                negocioId,
                omitirFechaHorario,
                buscarCercaDeMi,
                latitudUsuario,
                longitudUsuario,
                radioKm,
                usarUbigeo: false,
                usarNegocio: true,
                usarIgnorarHorario: true,
                usarFiltroDistancia: true);
        }
        catch (SqlException ex) when (ex.Message.Contains("@NegocioId", StringComparison.OrdinalIgnoreCase))
        {
            // Compatibilidad temporal mientras se despliega el filtro por negocio en BD.
            return await HomeBuscarEspaciosDisponiblesInternoAsync(
                fecha,
                horaInicio,
                horaFin,
                codigoDepartamento,
                codigoProvincia,
                codigoUbigeo,
                tipoDeporteId,
                null,
                omitirFechaHorario,
                buscarCercaDeMi,
                latitudUsuario,
                longitudUsuario,
                radioKm,
                usarUbigeo: true,
                usarNegocio: false,
                usarIgnorarHorario: true,
                usarFiltroDistancia: true);
        }
        catch (SqlException ex) when (ex.Message.Contains("@IgnorarFechaHorario", StringComparison.OrdinalIgnoreCase))
        {
            // Compatibilidad temporal mientras se despliega el nuevo parametro opcional de horario.
            return await HomeBuscarEspaciosDisponiblesInternoAsync(
                fecha,
                horaInicio,
                horaFin,
                codigoDepartamento,
                codigoProvincia,
                codigoUbigeo,
                tipoDeporteId,
                negocioId,
                omitirFechaHorario,
                buscarCercaDeMi,
                latitudUsuario,
                longitudUsuario,
                radioKm,
                usarUbigeo: true,
                usarNegocio: true,
                usarIgnorarHorario: false,
                usarFiltroDistancia: true);
        }
        catch (SqlException ex) when (
            ex.Message.Contains("@BuscarCercaDeMi", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("@LatitudUsuario", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("@LongitudUsuario", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("@RadioKm", StringComparison.OrdinalIgnoreCase))
        {
            // Compatibilidad temporal mientras se despliega el filtro de distancia en BD.
            return await HomeBuscarEspaciosDisponiblesInternoAsync(
                fecha,
                horaInicio,
                horaFin,
                codigoDepartamento,
                codigoProvincia,
                codigoUbigeo,
                tipoDeporteId,
                negocioId,
                omitirFechaHorario,
                buscarCercaDeMi,
                latitudUsuario,
                longitudUsuario,
                radioKm,
                usarUbigeo: true,
                usarNegocio: true,
                usarIgnorarHorario: true,
                usarFiltroDistancia: false);
        }
    }

    public async Task<(List<EspacioDisponibleViewModel> Espacios, int TotalRegistros)> HomeBuscarEspaciosDisponiblesPaginadoAsync(
        DateOnly fecha,
        TimeOnly horaInicio,
        TimeOnly horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId,
        int? negocioId,
        int pagina = 1,
        int tamanoPagina = 9,
        bool omitirFechaHorario = false,
        bool buscarCercaDeMi = false,
        decimal? latitudUsuario = null,
        decimal? longitudUsuario = null,
        decimal? radioKm = null)
    {
        var paginaEfectiva = pagina < 1 ? 1 : pagina;
        var tamanoEfectivo = tamanoPagina < 1 ? 9 : tamanoPagina;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_BuscarEspaciosDisponibles", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", horaInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", horaFin.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@CodigoDepartamento", string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoProvincia", string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim(), SqlDbType.Char);
        AddParam(cmd, "@CodigoUbigeo", string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim(), SqlDbType.Char);
        AddParam(cmd, "@TipoDeporteId", tipoDeporteId, SqlDbType.Int);
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@IgnorarFechaHorario", omitirFechaHorario, SqlDbType.Bit);
        AddParam(cmd, "@BuscarCercaDeMi", buscarCercaDeMi, SqlDbType.Bit);
        AddParam(cmd, "@LatitudUsuario", latitudUsuario, SqlDbType.Decimal);
        AddParam(cmd, "@LongitudUsuario", longitudUsuario, SqlDbType.Decimal);
        AddParam(cmd, "@RadioKm", radioKm, SqlDbType.Decimal);
        AddParam(cmd, "@Pagina", paginaEfectiva, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoEfectivo, SqlDbType.Int);
        var totalParam = new SqlParameter("@TotalRegistros", SqlDbType.Int) { Direction = ParameterDirection.Output };
        cmd.Parameters.Add(totalParam);

        var list = new List<EspacioDisponibleViewModel>();
        await using (var dr = await cmd.ExecuteReaderAsync())
        {
            while (await dr.ReadAsync())
            {
                if (dr.FieldCount >= 16)
                {
                    list.Add(new EspacioDisponibleViewModel
                    {
                        EspacioDeportivoId = dr.GetInt32(0),
                        NombreEspacio = dr.GetString(1),
                        Codigo = dr.GetString(2),
                        SedeNombre = dr.GetString(3),
                        SedeDireccion = !dr.IsDBNull(4) ? dr.GetString(4) : null,
                        SedeConsideracionesReserva = !dr.IsDBNull(5) ? dr.GetString(5) : null,
                        Departamento = !dr.IsDBNull(6) ? dr.GetString(6) : null,
                        Provincia = !dr.IsDBNull(7) ? dr.GetString(7) : null,
                        Distrito = !dr.IsDBNull(8) ? dr.GetString(8) : null,
                        TipoDeporteNombre = dr.GetString(9),
                        TipoSueloNombre = !dr.IsDBNull(10) ? dr.GetString(10) : null,
                        TarifaDesde = !dr.IsDBNull(11) ? dr.GetDecimal(11) : null,
                        TieneIluminacion = ReadBool(dr, 12),
                        Techada = ReadBool(dr, 13),
                        CorreoNotificacion = !dr.IsDBNull(14) ? dr.GetString(14) : null,
                        TelefonoContacto = dr.FieldCount > 21 ? (dr.IsDBNull(15) ? null : dr.GetString(15)) : null,
                        WhatsappContacto = dr.FieldCount > 21
                            ? (!dr.IsDBNull(16) ? dr.GetString(16) : null)
                            : (!dr.IsDBNull(15) ? dr.GetString(15) : null),
                        PermiteChatWhatsapp = dr.FieldCount > 21 ? ReadBool(dr, 17) : ReadBool(dr, 16),
                        SedeId = dr.FieldCount > 21
                            ? (dr.IsDBNull(18) ? null : dr.GetInt32(18))
                            : (dr.FieldCount > 17 && !dr.IsDBNull(17) ? dr.GetInt32(17) : null),
                        SedeMapaUrl = dr.FieldCount > 21
                            ? (dr.IsDBNull(19) ? null : dr.GetString(19))
                            : (dr.FieldCount > 20 && !dr.IsDBNull(18) ? dr.GetString(18) : null),
                        SedeFotoPrincipalUrl = dr.FieldCount > 21
                            ? (dr.IsDBNull(20) ? null : dr.GetString(20))
                            : (dr.FieldCount > 20
                                ? (dr.IsDBNull(19) ? null : dr.GetString(19))
                                : (dr.FieldCount > 18 && !dr.IsDBNull(18) ? dr.GetString(18) : null)),
                        SedeFotos = dr.FieldCount > 21
                            ? (dr.IsDBNull(21)
                                ? new List<string>()
                                : dr.GetString(21)
                                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                    .Where(x => !string.IsNullOrWhiteSpace(x))
                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                    .ToList())
                            : (dr.FieldCount > 19 && !dr.IsDBNull(19)
                                ? dr.GetString(19)
                                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                    .Where(x => !string.IsNullOrWhiteSpace(x))
                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                    .ToList()
                                : new List<string>()),
                        DistanciaKm = dr.FieldCount > 22 && !dr.IsDBNull(22) ? dr.GetDecimal(22) : null
                    });
                }
            }
        }

        var total = totalParam.Value is int t ? t : list.Count;
        return (list, total);
    }

    private async Task<List<EspacioDisponibleViewModel>> HomeBuscarEspaciosDisponiblesInternoAsync(
        DateOnly fecha,
        TimeOnly horaInicio,
        TimeOnly horaFin,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        int? tipoDeporteId,
        int? negocioId,
        bool omitirFechaHorario,
        bool buscarCercaDeMi,
        decimal? latitudUsuario,
        decimal? longitudUsuario,
        decimal? radioKm,
        bool usarUbigeo,
        bool usarNegocio,
        bool usarIgnorarHorario,
        bool usarFiltroDistancia)
    {
        var list = new List<EspacioDisponibleViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_BuscarEspaciosDisponibles", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", horaInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", horaFin.ToTimeSpan(), SqlDbType.Time);
        if (usarUbigeo)
        {
            AddParam(cmd, "@CodigoDepartamento", string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim(), SqlDbType.Char);
            AddParam(cmd, "@CodigoProvincia", string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim(), SqlDbType.Char);
            AddParam(cmd, "@CodigoUbigeo", string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim(), SqlDbType.Char);
        }
        else
        {
            AddParam(cmd, "@SedeId", null, SqlDbType.Int);
        }
        AddParam(cmd, "@TipoDeporteId", tipoDeporteId, SqlDbType.Int);
        if (usarNegocio)
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        if (usarIgnorarHorario)
            AddParam(cmd, "@IgnorarFechaHorario", omitirFechaHorario, SqlDbType.Bit);
        if (usarFiltroDistancia)
        {
            AddParam(cmd, "@BuscarCercaDeMi", buscarCercaDeMi, SqlDbType.Bit);
            AddParam(cmd, "@LatitudUsuario", latitudUsuario, SqlDbType.Decimal);
            AddParam(cmd, "@LongitudUsuario", longitudUsuario, SqlDbType.Decimal);
            AddParam(cmd, "@RadioKm", radioKm, SqlDbType.Decimal);
        }
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            if (dr.FieldCount >= 16)
            {
                list.Add(new EspacioDisponibleViewModel
                {
                    EspacioDeportivoId = dr.GetInt32(0),
                    NombreEspacio = dr.GetString(1),
                    Codigo = dr.GetString(2),
                    SedeNombre = dr.GetString(3),
                    SedeDireccion = !dr.IsDBNull(4) ? dr.GetString(4) : null,
                    SedeConsideracionesReserva = !dr.IsDBNull(5) ? dr.GetString(5) : null,
                    Departamento = !dr.IsDBNull(6) ? dr.GetString(6) : null,
                    Provincia = !dr.IsDBNull(7) ? dr.GetString(7) : null,
                    Distrito = !dr.IsDBNull(8) ? dr.GetString(8) : null,
                    TipoDeporteNombre = dr.GetString(9),
                    TipoSueloNombre = !dr.IsDBNull(10) ? dr.GetString(10) : null,
                    TarifaDesde = !dr.IsDBNull(11) ? dr.GetDecimal(11) : null,
                    TieneIluminacion = ReadBool(dr, 12),
                    Techada = ReadBool(dr, 13),
                    CorreoNotificacion = !dr.IsDBNull(14) ? dr.GetString(14) : null,
                    TelefonoContacto = dr.FieldCount > 21
                        ? (dr.IsDBNull(15) ? null : dr.GetString(15))
                        : null,
                    WhatsappContacto = dr.FieldCount > 21
                        ? (!dr.IsDBNull(16) ? dr.GetString(16) : null)
                        : (!dr.IsDBNull(15) ? dr.GetString(15) : null),
                    PermiteChatWhatsapp = dr.FieldCount > 21
                        ? ReadBool(dr, 17)
                        : ReadBool(dr, 16),
                    SedeId = dr.FieldCount > 21
                        ? (dr.IsDBNull(18) ? null : dr.GetInt32(18))
                        : (dr.FieldCount > 17 && !dr.IsDBNull(17) ? dr.GetInt32(17) : null),
                    SedeMapaUrl = dr.FieldCount > 21
                        ? (dr.IsDBNull(19) ? null : dr.GetString(19))
                        : (dr.FieldCount > 20 && !dr.IsDBNull(18) ? dr.GetString(18) : null),
                    SedeFotoPrincipalUrl = dr.FieldCount > 21
                        ? (dr.IsDBNull(20) ? null : dr.GetString(20))
                        : (dr.FieldCount > 20
                            ? (dr.IsDBNull(19) ? null : dr.GetString(19))
                            : (dr.FieldCount > 18 && !dr.IsDBNull(18) ? dr.GetString(18) : null)),
                    SedeFotos = dr.FieldCount > 21
                        ? (dr.IsDBNull(21)
                            ? new List<string>()
                            : dr.GetString(21)
                                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                .Where(x => !string.IsNullOrWhiteSpace(x))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList())
                        : (dr.FieldCount > 19 && !dr.IsDBNull(19)
                            ? dr.GetString(19)
                                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                .Where(x => !string.IsNullOrWhiteSpace(x))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList()
                            : new List<string>()),
                    DistanciaKm = dr.FieldCount > 22 && !dr.IsDBNull(22) ? dr.GetDecimal(22) : null
                });
            }
            else
            {
                list.Add(new EspacioDisponibleViewModel
                {
                    EspacioDeportivoId = dr.GetInt32(0),
                    NombreEspacio = dr.GetString(1),
                    Codigo = dr.GetString(2),
                    SedeNombre = dr.GetString(3),
                    SedeConsideracionesReserva = dr.FieldCount > 4 && !dr.IsDBNull(4) ? dr.GetString(4) : null,
                    TipoDeporteNombre = dr.FieldCount > 5 && !dr.IsDBNull(5) ? dr.GetString(5) : string.Empty,
                    TieneIluminacion = dr.FieldCount > 6 && ReadBool(dr, 6),
                    Techada = dr.FieldCount > 7 && ReadBool(dr, 7),
                    WhatsappContacto = dr.FieldCount > 8 && !dr.IsDBNull(8) ? dr.GetString(8) : null,
                    PermiteChatWhatsapp = dr.FieldCount > 9 && ReadBool(dr, 9),
                    DistanciaKm = null
                });
            }
        }
        return list;
    }

    public async Task<int> HomeSolicitarReservaPublicaAsync(SolicitudReservaPublicaFormViewModel model)
    {
        try
        {
            return await HomeSolicitarReservaPublicaInternoAsync(model, incluirUsuarioId: true);
        }
        catch (SqlException ex) when (ex.Message.Contains("@UsuarioId", StringComparison.OrdinalIgnoreCase))
        {
            return await HomeSolicitarReservaPublicaInternoAsync(model, incluirUsuarioId: false);
        }
    }

    private async Task<int> HomeSolicitarReservaPublicaInternoAsync(SolicitudReservaPublicaFormViewModel model, bool incluirUsuarioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Home_SolicitarReservaPublica", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", model.Fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", model.HoraInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", model.HoraFin.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@Nombres", model.Nombres, SqlDbType.NVarChar);
        AddParam(cmd, "@Apellidos", model.Apellidos, SqlDbType.NVarChar);
        AddParam(cmd, "@NombreEquipo", model.NombreEquipo, SqlDbType.NVarChar);
        AddParam(cmd, "@TipoDocumento", model.TipoDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@NumeroDocumento", model.NumeroDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@Telefono", model.Telefono, SqlDbType.NVarChar);
        AddParam(cmd, "@Correo", model.Correo, SqlDbType.NVarChar);
        AddParam(cmd, "@Comentario", model.Comentario, SqlDbType.NVarChar);
        AddParam(cmd, "@CodigoCupon", string.IsNullOrWhiteSpace(model.CodigoCupon) ? null : model.CodigoCupon.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
        if (incluirUsuarioId)
            AddParam(cmd, "@UsuarioId", string.IsNullOrWhiteSpace(model.UsuarioId) ? null : model.UsuarioId.Trim(), SqlDbType.NVarChar);
        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result);
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

    public async Task<(int TotalSedes, int TotalEspacios, int ReservasHoy, decimal IngresosHoy, decimal OcupacionHoyPct, int NoShowMes, decimal TicketPromedioMes)> PanelObtenerMetricasAsync(int negocioId, DateOnly fecha, int? sedeId = null)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Panel_ObtenerMetricas", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
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
