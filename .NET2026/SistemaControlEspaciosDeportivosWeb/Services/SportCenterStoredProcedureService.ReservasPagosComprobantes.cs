using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<(List<ReservaItemViewModel> Reservas, int TotalRegistros)> ReservasListarAsync(int negocioId, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int? sedeId = null, int? espacioDeportivoId = null, int? estado = null, string? estadosCsv = null, int pagina = 1, int tamanoPagina = 20)
    {
        var list = new List<ReservaItemViewModel>();
        var paginaNormalizada = pagina < 1 ? 1 : pagina;
        var tamanoNormalizado = tamanoPagina < 1 ? 20 : tamanoPagina;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Estado", estado, SqlDbType.Int);
        AddParam(cmd, "@EstadosCsv", string.IsNullOrWhiteSpace(estadosCsv) ? null : estadosCsv, SqlDbType.NVarChar);
        AddParam(cmd, "@Pagina", paginaNormalizada, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoNormalizado, SqlDbType.Int);
        var totalRegistrosParam = cmd.Parameters.Add("@TotalRegistros", SqlDbType.Int);
        totalRegistrosParam.Direction = ParameterDirection.Output;
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReservaItemViewModel
            {
                Id = dr.GetInt32(0),
                Cliente = dr.GetString(1),
                Equipo = dr.FieldCount > 2 && !dr.IsDBNull(2) ? dr.GetString(2) : null,
                Espacio = dr.GetString(3),
                Sede = dr.GetString(4),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(5)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(6)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(7)),
                Total = dr.GetDecimal(8),
                Adelanto = dr.FieldCount > 9 && !dr.IsDBNull(9) ? dr.GetDecimal(9) : 0m,
                SaldoPendiente = dr.FieldCount > 10 && !dr.IsDBNull(10) ? dr.GetDecimal(10) : dr.GetDecimal(8),
                Estado = dr.GetString(dr.FieldCount > 11 ? 11 : 9)
            });
        }
        await dr.CloseAsync();
        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        return (list, totalRegistros);
    }

    public async Task<ReservasListadoResumenViewModel> ReservasListadoResumenAsync(int negocioId, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int? sedeId = null, int? espacioDeportivoId = null, int? estado = null, string? estadosCsv = null)
    {
        var resumen = new ReservasListadoResumenViewModel();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_ListadoResumen", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Estado", estado, SqlDbType.Int);
        AddParam(cmd, "@EstadosCsv", string.IsNullOrWhiteSpace(estadosCsv) ? null : estadosCsv, SqlDbType.NVarChar);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (await dr.ReadAsync())
        {
            // Compatibilidad temporal:
            // - Version nueva SP: [TotalReservasActivas, TotalPendientes, TotalPagadas, SaldoTotal]
            // - Version anterior SP: [TotalPendientes, TotalPagadas, SaldoTotal]
            if (dr.FieldCount >= 4)
            {
                resumen.TotalReservasActivas = dr.IsDBNull(0) ? 0 : Convert.ToInt32(dr.GetValue(0));
                resumen.TotalPendientes = dr.IsDBNull(1) ? 0 : Convert.ToInt32(dr.GetValue(1));
                resumen.TotalPagadas = dr.IsDBNull(2) ? 0 : Convert.ToInt32(dr.GetValue(2));
                resumen.SaldoTotal = dr.IsDBNull(3) ? 0m : Convert.ToDecimal(dr.GetValue(3));
            }
            else
            {
                resumen.TotalPendientes = dr.IsDBNull(0) ? 0 : Convert.ToInt32(dr.GetValue(0));
                resumen.TotalPagadas = dr.IsDBNull(1) ? 0 : Convert.ToInt32(dr.GetValue(1));
                resumen.SaldoTotal = dr.IsDBNull(2) ? 0m : Convert.ToDecimal(dr.GetValue(2));
                resumen.TotalReservasActivas = resumen.TotalPendientes + resumen.TotalPagadas;
            }
        }

        return resumen;
    }

    public async Task<ReservaFormViewModel?> ReservasObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        return new ReservaFormViewModel
        {
            Id = dr.GetInt32(0),
            EspacioDeportivoId = dr.GetInt32(1),
            ClienteId = dr.GetInt32(2),
            Fecha = DateOnly.FromDateTime(dr.GetDateTime(3)),
            HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(4)),
            HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(5)),
            Total = dr.GetDecimal(6),
            Adelanto = dr.GetDecimal(7),
            Estado = (EstadoReserva)dr.GetInt32(8),
            Comentario = dr.FieldCount > 9 && !dr.IsDBNull(9) ? dr.GetString(9) : null,
            NegocioId = negocioId
        };
    }

    public async Task<int> ReservasCrearAsync(ReservaFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@ClienteId", model.ClienteId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", model.Fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", model.HoraInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", model.HoraFin.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@Total", model.Total, SqlDbType.Decimal);
        AddParam(cmd, "@Adelanto", model.Adelanto, SqlDbType.Decimal);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@RegistrarPago", model.RegistrarPago, SqlDbType.Bit);
        AddParam(cmd, "@FormaPagoId", model.FormaPagoId, SqlDbType.Int);
        AddParam(cmd, "@FechaPago", model.FechaPago, SqlDbType.DateTime2);
        AddParam(cmd, "@NumeroOperacion", model.NumeroOperacion, SqlDbType.NVarChar);
        AddParam(cmd, "@Comentario", model.Comentario, SqlDbType.NVarChar);
        AddParam(cmd, "@CanalOrigen", "ADMIN", SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> ReservasActualizarAsync(ReservaFormViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Reservas_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
            AddParam(cmd, "@ClienteId", model.ClienteId, SqlDbType.Int);
            AddParam(cmd, "@Fecha", model.Fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@HoraInicio", model.HoraInicio.ToTimeSpan(), SqlDbType.Time);
            AddParam(cmd, "@HoraFin", model.HoraFin.ToTimeSpan(), SqlDbType.Time);
            AddParam(cmd, "@Total", model.Total, SqlDbType.Decimal);
            AddParam(cmd, "@Adelanto", model.Adelanto, SqlDbType.Decimal);
            AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
            AddParam(cmd, "@RegistrarPago", model.RegistrarPago, SqlDbType.Bit);
            AddParam(cmd, "@FormaPagoId", model.FormaPagoId, SqlDbType.Int);
            AddParam(cmd, "@FechaPago", model.FechaPago, SqlDbType.DateTime2);
            AddParam(cmd, "@NumeroOperacion", model.NumeroOperacion, SqlDbType.NVarChar);
            AddParam(cmd, "@Comentario", model.Comentario, SqlDbType.NVarChar);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> ReservasEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Reservas_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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

    public async Task<bool> ReservasCambiarEstadoRapidoAsync(int negocioId, int id, int nuevoEstado, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Reservas_CambiarEstadoRapido", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@NuevoEstado", nuevoEstado, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<List<ReservaHistorialItemViewModel>> ReservasHistorialAsync(int negocioId, int reservaId)
    {
        var list = new List<ReservaHistorialItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_Historial", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@ReservaId", reservaId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReservaHistorialItemViewModel
            {
                FechaRegistro = dr.GetDateTime(0),
                Accion = dr.IsDBNull(1) ? string.Empty : dr.GetString(1),
                Usuario = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
                Detalle = dr.IsDBNull(3) ? null : dr.GetString(3)
            });
        }

        return list;
    }

    public async Task<ReservaRecordatorioPendienteViewModel?> ReservasObtenerParaRecordatorioAsync(int negocioId, int reservaId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_ObtenerParaRecordatorio", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@ReservaId", reservaId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;

        return new ReservaRecordatorioPendienteViewModel
        {
            ReservaId = dr.GetInt32(0),
            NegocioId = dr.GetInt32(1),
            Cliente = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
            Correo = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
            Sede = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
            Espacio = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
            Fecha = DateOnly.FromDateTime(dr.GetDateTime(6)),
            HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(7)),
            HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(8)),
            CorreoNotificacion = dr.IsDBNull(9) ? null : dr.GetString(9),
            WhatsappContacto = dr.IsDBNull(10) ? null : dr.GetString(10)
        };
    }

    public async Task<List<ReservaCalendarioEventoViewModel>> ReservasCalendarioEventosAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null, int? espacioDeportivoId = null, int? estado = null)
    {
        var list = new List<ReservaCalendarioEventoViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_CalendarioEventos", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Estado", estado, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReservaCalendarioEventoViewModel
            {
                Id = dr.GetInt32(0),
                TipoEvento = dr.GetString(1),
                Titulo = dr.GetString(2),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(3)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(4)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(5)),
                Estado = dr.IsDBNull(6) ? null : dr.GetInt32(6),
                Color = dr.IsDBNull(7) ? null : dr.GetString(7),
                EspacioDeportivoId = dr.IsDBNull(8) ? null : dr.GetInt32(8),
                Espacio = dr.IsDBNull(9) ? string.Empty : dr.GetString(9),
                Sede = dr.IsDBNull(10) ? string.Empty : dr.GetString(10),
                Motivo = dr.FieldCount > 11 && !dr.IsDBNull(11) ? dr.GetString(11) : null,
                EstadoCodigo = dr.FieldCount > 12 && !dr.IsDBNull(12) ? dr.GetString(12) : null,
                EstadoTexto = dr.FieldCount > 13 && !dr.IsDBNull(13) ? dr.GetString(13) : null,
                TotalReserva = dr.FieldCount > 14 && !dr.IsDBNull(14) ? dr.GetDecimal(14) : 0m
            });
        }
        return list;
    }

    public async Task<bool> ReservasMoverAsync(int negocioId, int id, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Reservas_Mover", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Id", id, SqlDbType.Int);
            AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@HoraInicio", horaInicio.ToTimeSpan(), SqlDbType.Time);
            AddParam(cmd, "@HoraFin", horaFin.ToTimeSpan(), SqlDbType.Time);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<ReservaDisponibilidadValidacionViewModel> ReservasValidarDisponibilidadAsync(int negocioId, int? reservaId, int espacioDeportivoId, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_ValidarDisponibilidad", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@ReservaId", reservaId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", horaInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", horaFin.ToTimeSpan(), SqlDbType.Time);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
        {
            return new ReservaDisponibilidadValidacionViewModel
            {
                Disponible = false,
                Mensaje = "No se pudo validar disponibilidad."
            };
        }

        var disponible = ReadBool(dr, 0);
        var mensaje = dr.IsDBNull(1) ? (disponible ? "Disponible." : "Horario no disponible.") : dr.GetString(1);

        string? conflictoTipo = null;
        int? conflictoId = null;
        if (dr.FieldCount > 2 && !dr.IsDBNull(2))
        {
            conflictoTipo = dr.GetString(2);
        }
        if (dr.FieldCount > 3 && !dr.IsDBNull(3))
        {
            conflictoId = dr.GetInt32(3);
        }

        return new ReservaDisponibilidadValidacionViewModel
        {
            Disponible = disponible,
            Mensaje = mensaje,
            ConflictoTipo = conflictoTipo,
            ConflictoId = conflictoId
        };
    }

    public async Task<ReservaCotizacionViewModel> ReservasCotizarAsync(int negocioId, int espacioDeportivoId, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_Cotizar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", horaInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", horaFin.ToTimeSpan(), SqlDbType.Time);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
        {
            return new ReservaCotizacionViewModel
            {
                Ok = false,
                Mensaje = "No se pudo calcular la tarifa."
            };
        }

        return new ReservaCotizacionViewModel
        {
            Ok = true,
            Mensaje = dr.IsDBNull(0) ? string.Empty : dr.GetString(0),
            PrecioBase = dr.IsDBNull(1) ? 0m : dr.GetDecimal(1),
            DescuentoPct = dr.IsDBNull(2) ? 0m : dr.GetDecimal(2),
            PrecioFinal = dr.IsDBNull(3) ? 0m : dr.GetDecimal(3),
            MonedaSimbolo = dr.IsDBNull(4) ? "S/" : dr.GetString(4),
            MonedaNombre = dr.IsDBNull(5) ? "PEN" : dr.GetString(5),
            PoliticaConfirmacionPago = dr.IsDBNull(6) ? 0 : dr.GetByte(6),
            PorcentajeAdelantoMinimo = dr.IsDBNull(7) ? null : dr.GetDecimal(7)
        };
    }

    public async Task<List<BloqueoHorarioItemViewModel>> BloqueosListarAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null, int? espacioDeportivoId = null)
    {
        var list = new List<BloqueoHorarioItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Bloqueos_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new BloqueoHorarioItemViewModel
            {
                Id = dr.GetInt32(0),
                Sede = dr.GetString(1),
                Espacio = dr.GetString(2),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(3)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(4)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(5)),
                Motivo = dr.GetString(6),
                Activo = dr.GetBoolean(7)
            });
        }
        return list;
    }

    public async Task<int> BloqueosCrearAsync(BloqueoHorarioFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Bloqueos_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", model.EspacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Fecha", model.Fecha.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@HoraInicio", model.HoraInicio.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@HoraFin", model.HoraFin.ToTimeSpan(), SqlDbType.Time);
        AddParam(cmd, "@Motivo", model.Motivo, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> BloqueosEliminarAsync(int negocioId, int id, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Bloqueos_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        await cmd.ExecuteNonQueryAsync();
        return true;
    }

    public Task<List<SelectListItem>> ReservasComboEspaciosAsync(int negocioId, int? sedeId = null) => ComboAsync("Sp_Combos_EspaciosPorNegocio", ("@NegocioId", (object?)negocioId, SqlDbType.Int), ("@SedeId", sedeId, SqlDbType.Int));
    public Task<List<SelectListItem>> ReservasComboClientesAsync(int negocioId) => ComboAsync("Sp_Combos_Clientes", ("@NegocioId", (object?)negocioId, SqlDbType.Int));
    public Task<List<SelectListItem>> ReservasBuscarClientesAsync(int negocioId, string? buscar = null, int? clienteId = null, int top = 50) => ComboAsync(
        "Sp_Combos_Clientes_Buscar",
        ("@NegocioId", (object?)negocioId, SqlDbType.Int),
        ("@Buscar", string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(), SqlDbType.NVarChar),
        ("@ClienteId", clienteId, SqlDbType.Int),
        ("@Top", top < 1 ? 50 : top, SqlDbType.Int));

    public async Task<(List<PagoReservaResumenViewModel> Pagos, int TotalRegistros)> PagosListarAsync(int negocioId, int? sedeId = null, string? buscar = null, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int pagina = 1, int tamanoPagina = 20)
    {
        try
        {
            return await PagosListarInternoAsync(negocioId, sedeId, buscar, fechaDesde, fechaHasta, pagina, tamanoPagina, incluirRangoFecha: true);
        }
        catch (SqlException ex) when (ex.Message.Contains("@FechaDesde", StringComparison.OrdinalIgnoreCase) || ex.Message.Contains("@FechaHasta", StringComparison.OrdinalIgnoreCase))
        {
            // Compatibilidad con SP antiguo sin parametros de rango de fecha.
            return await PagosListarInternoAsync(negocioId, sedeId, buscar, null, null, pagina, tamanoPagina, incluirRangoFecha: false);
        }
    }

    public async Task<PagoReservaEditViewModel?> PagosObtenerAsync(int negocioId, int reservaId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", reservaId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        var vm = new PagoReservaEditViewModel
        {
            NegocioId = negocioId,
            ReservaId = dr.GetInt32(0),
            ReservaCodigo = dr.IsDBNull(1) ? $"#{dr.GetInt32(0)}" : dr.GetString(1),
            Sede = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
            Espacio = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
            Cliente = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
            FechaReserva = DateOnly.FromDateTime(dr.GetDateTime(5)),
            HoraInicioReserva = dr.IsDBNull(6) ? TimeOnly.MinValue : TimeOnly.FromTimeSpan(dr.GetTimeSpan(6)),
            HoraFinReserva = dr.IsDBNull(7) ? TimeOnly.MinValue : TimeOnly.FromTimeSpan(dr.GetTimeSpan(7)),
            TotalReserva = dr.IsDBNull(8) ? 0m : dr.GetDecimal(8),
            TotalPagado = dr.IsDBNull(9) ? 0m : dr.GetDecimal(9),
            SaldoPendiente = dr.IsDBNull(10) ? 0m : dr.GetDecimal(10),
            MonedaSimbolo = dr.IsDBNull(11) ? "S/" : dr.GetString(11),
            PoliticaConfirmacionPago = dr.IsDBNull(12) ? 0 : dr.GetInt32(12),
            PorcentajeAdelantoMinimo = dr.IsDBNull(13) ? null : dr.GetDecimal(13),
            TieneComprobanteActivo = dr.FieldCount > 14 && !dr.IsDBNull(14) && dr.GetBoolean(14),
            ReferenciaComprobante = dr.FieldCount > 15 && !dr.IsDBNull(15) ? dr.GetString(15) : string.Empty,
            NuevaFechaPago = DateTime.Today
        };

        if (await dr.NextResultAsync())
        {
            while (await dr.ReadAsync())
            {
                vm.Pagos.Add(new PagoReservaDetalleItemViewModel
                {
                    PagoId = dr.GetInt32(0),
                    FechaPago = dr.GetDateTime(1),
                    Monto = dr.GetDecimal(2),
                    FormaPagoId = dr.GetInt32(3),
                    FormaPagoNombre = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
                    NumeroOperacion = dr.IsDBNull(5) ? null : dr.GetString(5),
                    Observacion = dr.IsDBNull(6) ? null : dr.GetString(6)
                });
            }
        }

        return vm;
    }

    public async Task<int> PagosCrearAsync(PagoFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@ReservaId", model.ReservaId, SqlDbType.Int);
        AddParam(cmd, "@FechaPago", model.FechaPago, SqlDbType.DateTime2);
        AddParam(cmd, "@Monto", model.Monto, SqlDbType.Decimal);
        AddParam(cmd, "@FormaPago", model.FormaPagoId, SqlDbType.Int);
        AddParam(cmd, "@NumeroOperacion", model.NumeroOperacion, SqlDbType.NVarChar);
        AddParam(cmd, "@Observacion", model.Observacion, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> PagosActualizarAsync(int negocioId, int pagoId, string? observacion, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Pagos_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", pagoId, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@Observacion", string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim(), SqlDbType.NVarChar);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> PagosEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Pagos_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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

    public async Task<bool> PagosEliminarPorReservaAsync(int negocioId, int reservaId, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Pagos_EliminarPorReserva", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            AddParam(cmd, "@ReservaId", reservaId, SqlDbType.Int);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public Task<List<SelectListItem>> PagosComboReservasAsync(int negocioId, int? sedeId = null) => ComboAsync("Sp_Combos_ReservasPorNegocio", ("@NegocioId", (object?)negocioId, SqlDbType.Int), ("@SedeId", sedeId, SqlDbType.Int));
    public Task<List<SelectListItem>> PagosBuscarReservasAsync(int negocioId, string? buscar = null, int? reservaId = null, int top = 30) => ComboAsync(
        "Sp_Combos_Reservas_Buscar",
        ("@NegocioId", (object?)negocioId, SqlDbType.Int),
        ("@Buscar", string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(), SqlDbType.NVarChar),
        ("@ReservaId", reservaId, SqlDbType.Int),
        ("@Top", top < 1 ? 30 : top, SqlDbType.Int));
    public Task<List<SelectListItem>> PagosComboFormasPagoAsync(int negocioId) => ComboAsync("Sp_Combos_FormasPago", ("@NegocioId", negocioId, SqlDbType.Int));

    public async Task<(List<ComprobanteItemViewModel> Comprobantes, int TotalRegistros)> ComprobantesListarAsync(int negocioId, int? sedeId = null, string? buscar = null, string? codigoDocumento = null, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int pagina = 1, int tamanoPagina = 20)
    {
        try
        {
            return await ComprobantesListarInternoAsync(negocioId, sedeId, buscar, codigoDocumento, fechaDesde, fechaHasta, pagina, tamanoPagina, incluirRangoFecha: true);
        }
        catch (SqlException ex) when (ex.Message.Contains("@FechaDesde", StringComparison.OrdinalIgnoreCase) || ex.Message.Contains("@FechaHasta", StringComparison.OrdinalIgnoreCase))
        {
            // Compatibilidad con SP antiguo sin parametros de rango de fecha.
            return await ComprobantesListarInternoAsync(negocioId, sedeId, buscar, codigoDocumento, null, null, pagina, tamanoPagina, incluirRangoFecha: false);
        }
    }

    private async Task<(List<PagoReservaResumenViewModel> Pagos, int TotalRegistros)> PagosListarInternoAsync(int negocioId, int? sedeId, string? buscar, DateOnly? fechaDesde, DateOnly? fechaHasta, int pagina, int tamanoPagina, bool incluirRangoFecha)
    {
        var list = new List<PagoReservaResumenViewModel>();
        var paginaNormalizada = pagina < 1 ? 1 : pagina;
        var tamanoNormalizado = tamanoPagina < 1 ? 20 : tamanoPagina;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@Buscar", string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(), SqlDbType.NVarChar);
        if (incluirRangoFecha)
        {
            AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        }
        AddParam(cmd, "@Pagina", paginaNormalizada, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoNormalizado, SqlDbType.Int);
        var totalRegistrosParam = cmd.Parameters.Add("@TotalRegistros", SqlDbType.Int);
        totalRegistrosParam.Direction = ParameterDirection.Output;
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new PagoReservaResumenViewModel
            {
                ReservaId = dr.GetInt32(0),
                ReservaCodigo = dr.IsDBNull(1) ? $"#{dr.GetInt32(0)}" : dr.GetString(1),
                Sede = dr.IsDBNull(2) ? string.Empty : dr.GetString(2),
                Espacio = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                Cliente = dr.IsDBNull(4) ? string.Empty : dr.GetString(4),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(5)),
                MontoTotal = dr.IsDBNull(6) ? 0m : dr.GetDecimal(6),
                SaldoPendiente = dr.IsDBNull(7) ? 0m : dr.GetDecimal(7),
                FormaPagoResumen = dr.IsDBNull(8) ? string.Empty : dr.GetString(8),
                CantidadPagos = dr.IsDBNull(9) ? 0 : dr.GetInt32(9),
                MonedaSimbolo = dr.IsDBNull(10) ? "S/" : dr.GetString(10),
                PagadaCompleta = dr.FieldCount > 11
                    ? !dr.IsDBNull(11) && dr.GetBoolean(11)
                    : (dr.IsDBNull(7) || dr.GetDecimal(7) <= 0m),
                TieneComprobanteActivo = dr.FieldCount > 12
                    ? !dr.IsDBNull(12) && dr.GetBoolean(12)
                    : false,
                Referencia = dr.FieldCount > 13 && !dr.IsDBNull(13)
                    ? dr.GetString(13)
                    : string.Empty
            });
        }
        await dr.CloseAsync();
        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        return (list, totalRegistros);
    }

    private async Task<(List<ComprobanteItemViewModel> Comprobantes, int TotalRegistros)> ComprobantesListarInternoAsync(int negocioId, int? sedeId, string? buscar, string? codigoDocumento, DateOnly? fechaDesde, DateOnly? fechaHasta, int pagina, int tamanoPagina, bool incluirRangoFecha)
    {
        var list = new List<ComprobanteItemViewModel>();
        var paginaNormalizada = pagina < 1 ? 1 : pagina;
        var tamanoNormalizado = tamanoPagina < 1 ? 20 : tamanoPagina;
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@Buscar", string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(), SqlDbType.NVarChar);
        AddParam(cmd, "@CodigoDocumento", string.IsNullOrWhiteSpace(codigoDocumento) ? null : codigoDocumento.Trim().ToUpperInvariant(), SqlDbType.NVarChar);
        if (incluirRangoFecha)
        {
            AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
            AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        }
        AddParam(cmd, "@Pagina", paginaNormalizada, SqlDbType.Int);
        AddParam(cmd, "@TamanoPagina", tamanoNormalizado, SqlDbType.Int);
        var totalRegistrosParam = cmd.Parameters.Add("@TotalRegistros", SqlDbType.Int);
        totalRegistrosParam.Direction = ParameterDirection.Output;
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ComprobanteItemViewModel
            {
                Id = dr.GetInt32(0),
                Tipo = dr.GetString(1),
                SerieNumero = dr.GetString(2),
                FechaEmision = dr.GetDateTime(3),
                Cliente = dr.GetString(4),
                Total = dr.GetDecimal(5),
                Estado = dr.GetString(6),
                EstadoCodigo = dr.FieldCount > 7 && !dr.IsDBNull(7) ? dr.GetInt32(7) : 0,
                CodigoDocumentoComprobante = dr.FieldCount > 8 && !dr.IsDBNull(8) ? dr.GetString(8) : string.Empty,
                Referencia = dr.FieldCount > 9 && !dr.IsDBNull(9) ? dr.GetString(9) : "-",
                TieneNotasRelacionadas = dr.FieldCount > 10 && !dr.IsDBNull(10) && dr.GetBoolean(10),
                ReservaId = dr.IsDBNull(11) ? 0 : dr.GetInt32(11),
                EsTributario = dr.FieldCount > 12 && !dr.IsDBNull(12) && dr.GetBoolean(12),
                UrlDescargaProveedor = dr.FieldCount > 13 && !dr.IsDBNull(13) ? dr.GetString(13) : null
            });
        }
        await dr.CloseAsync();
        var totalRegistros = totalRegistrosParam.Value is int total ? total : 0;
        return (list, totalRegistros);
    }

    public async Task<ComprobanteVisualizacionViewModel?> ComprobantesObtenerVisualizacionAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_ObtenerVisualizacion", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        return new ComprobanteVisualizacionViewModel
        {
            Id = dr.GetInt32(0),
            NegocioId = dr.GetInt32(1),
            ReservaId = dr.GetInt32(2),
            TipoComprobante = dr.GetInt32(3),
            CodigoDocumentoComprobante = dr.GetString(4),
            TipoDocumentoNombre = dr.GetString(5),
            EsTributario = dr.GetBoolean(6),
            Serie = dr.GetString(7),
            Numero = dr.GetInt32(8),
            FechaEmision = dr.GetDateTime(9),
            MonedaSimbolo = dr.IsDBNull(10) ? "S/" : dr.GetString(10),
            SubTotal = dr.GetDecimal(11),
            Igv = dr.GetDecimal(12),
            Total = dr.GetDecimal(13),
            PorcentajeIgv = dr.IsDBNull(14) ? 0 : dr.GetInt32(14),
            NegocioNombre = dr.IsDBNull(15) ? string.Empty : dr.GetString(15),
            NegocioRazonSocial = dr.IsDBNull(16) ? null : dr.GetString(16),
            NegocioDireccionFiscal = dr.IsDBNull(17) ? null : dr.GetString(17),
            NegocioDistrito = dr.IsDBNull(18) ? null : dr.GetString(18),
            NegocioProvincia = dr.IsDBNull(19) ? null : dr.GetString(19),
            NegocioDepartamento = dr.IsDBNull(20) ? null : dr.GetString(20),
            NegocioDocumento = dr.IsDBNull(21) ? null : dr.GetString(21),
            ClienteNombre = dr.IsDBNull(22) ? string.Empty : dr.GetString(22),
            ClienteDocumento = dr.IsDBNull(23) ? null : dr.GetString(23),
            ClienteDireccion = dr.IsDBNull(24) ? null : dr.GetString(24),
            ClienteDistrito = dr.IsDBNull(25) ? null : dr.GetString(25),
            ClienteProvincia = dr.IsDBNull(26) ? null : dr.GetString(26),
            ClienteDepartamento = dr.IsDBNull(27) ? null : dr.GetString(27),
            ClienteCorreo = dr.IsDBNull(28) ? null : dr.GetString(28),
            SedeNombre = dr.IsDBNull(29) ? string.Empty : dr.GetString(29),
            EspacioNombre = dr.IsDBNull(30) ? string.Empty : dr.GetString(30),
            FechaReserva = DateOnly.FromDateTime(dr.GetDateTime(31)),
            HoraInicioReserva = TimeOnly.FromTimeSpan(dr.GetTimeSpan(32)),
            HoraFinReserva = TimeOnly.FromTimeSpan(dr.GetTimeSpan(33)),
            UrlDescargaProveedor = dr.IsDBNull(34) ? null : dr.GetString(34)
        };
    }

    public async Task<ComprobanteFormViewModel?> ComprobantesObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        return new ComprobanteFormViewModel
        {
            Id = dr.GetInt32(0),
            ReservaId = dr.GetInt32(1),
            TipoComprobante = (TipoComprobante)dr.GetInt32(2),
            Serie = dr.GetString(3),
            Numero = dr.GetInt32(4),
            FechaEmision = dr.GetDateTime(5),
            TipoMoneda = (TipoMoneda)dr.GetInt32(6),
            SubTotal = dr.GetDecimal(7),
            Igv = dr.GetDecimal(8),
            Total = dr.GetDecimal(9),
            Estado = (EstadoComprobanteElectronico)dr.GetInt32(10),
            CodigoDocumentoComprobante = dr.FieldCount > 11 && !dr.IsDBNull(11) ? dr.GetString(11) : "03",
            ComprobanteReferenciaId = dr.FieldCount > 12 && !dr.IsDBNull(12) ? dr.GetInt32(12) : null,
            TipoNota = dr.FieldCount > 13 && !dr.IsDBNull(13) ? dr.GetString(13) : null,
            TipoNotaCodigoSunat = dr.FieldCount > 14 && !dr.IsDBNull(14) ? dr.GetString(14) : null,
            NegocioId = negocioId
        };
    }

    public async Task<int> ComprobantesCrearAsync(ComprobanteFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_Crear", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@ReservaId", model.ReservaId, SqlDbType.Int);
        AddParam(cmd, "@TipoComprobante", (int)model.TipoComprobante, SqlDbType.Int);
        AddParam(cmd, "@CodigoDocumentoComprobante", model.CodigoDocumentoComprobante, SqlDbType.NVarChar);
        AddParam(cmd, "@NegocioSerieId", model.NegocioSerieId, SqlDbType.Int);
        AddParam(cmd, "@Serie", model.Serie, SqlDbType.NVarChar);
        AddParam(cmd, "@Numero", model.Numero > 0 ? model.Numero : null, SqlDbType.Int);
        AddParam(cmd, "@FechaEmision", model.FechaEmision, SqlDbType.DateTime2);
        AddParam(cmd, "@TipoMoneda", (int)model.TipoMoneda, SqlDbType.Int);
        AddParam(cmd, "@SubTotal", model.SubTotal, SqlDbType.Decimal);
        AddParam(cmd, "@Igv", model.Igv, SqlDbType.Decimal);
        AddParam(cmd, "@Total", model.Total, SqlDbType.Decimal);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
        AddParam(cmd, "@ClienteCorreo", model.ClienteCorreo, SqlDbType.NVarChar);
        AddParam(cmd, "@ClienteTipoDocumento", model.ClienteTipoDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@ClienteNumeroDocumento", model.ClienteNumeroDocumento, SqlDbType.NVarChar);
        AddParam(cmd, "@ClienteDireccionFiscal", model.ClienteDireccionFiscal, SqlDbType.NVarChar);
        AddParam(cmd, "@ClienteCodigoUbigeo", model.ClienteCodigoUbigeo, SqlDbType.Char);
        AddParam(cmd, "@ComprobanteReferenciaId", model.ComprobanteReferenciaId, SqlDbType.Int);
        AddParam(cmd, "@TipoNota", model.TipoNota, SqlDbType.Char);
        AddParam(cmd, "@TipoNotaCodigoSunat", model.TipoNotaCodigoSunat, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> ComprobantesActualizarAsync(ComprobanteFormViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Comprobantes_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@ClienteCorreo", model.ClienteCorreo, SqlDbType.NVarChar);
            AddParam(cmd, "@ClienteTipoDocumento", model.ClienteTipoDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@ClienteNumeroDocumento", model.ClienteNumeroDocumento, SqlDbType.NVarChar);
            AddParam(cmd, "@ClienteDireccionFiscal", model.ClienteDireccionFiscal, SqlDbType.NVarChar);
            AddParam(cmd, "@ClienteCodigoUbigeo", model.ClienteCodigoUbigeo, SqlDbType.Char);
            AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
            await cmd.ExecuteNonQueryAsync();
            return true;
        }
        catch (SqlException ex) when (EsErrorNoEncontrado(ex.Message))
        {
            return false;
        }
    }

    public async Task<bool> ComprobantesEliminarAsync(int negocioId, int id, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Comprobantes_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
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

    public Task<List<SelectListItem>> ComprobantesComboReservasAsync(int negocioId, int? sedeId = null) => ComboAsync("Sp_Combos_ReservasPorNegocio", ("@NegocioId", (object?)negocioId, SqlDbType.Int), ("@SedeId", sedeId, SqlDbType.Int));

    public Task<List<SelectListItem>> ComprobantesBuscarReservasPagadasAsync(int negocioId, string? buscar = null, int? reservaId = null, int top = 30) => ComboAsync(
        "Sp_Combos_ReservasPagadas_Buscar",
        ("@NegocioId", (object?)negocioId, SqlDbType.Int),
        ("@Buscar", string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim(), SqlDbType.NVarChar),
        ("@ReservaId", reservaId, SqlDbType.Int),
        ("@Top", top < 1 ? 30 : top, SqlDbType.Int));

    public async Task<ComprobanteReservaContextoViewModel?> ComprobantesObtenerContextoReservaAsync(int negocioId, int reservaId, string? codigoSunat = null)
    {
        var pago = await PagosObtenerAsync(negocioId, reservaId);
        if (pago is null) return null;

        int? sedeId = null;
        int? clienteId = null;
        string? clienteCorreo = null;
        string? clienteTipoDocumento = null;
        string? clienteNumeroDocumento = null;
        string? clienteDireccionFiscal = null;
        string? clienteCodigoUbigeo = null;
        string? clienteCodigoDepartamento = null;
        string? clienteCodigoProvincia = null;
        await using (var cn = CreateConnection())
        {
            await cn.OpenAsync();
            await using var cmd = new SqlCommand(@"
                SELECT TOP (1)
                    e.SedeId,
                    c.Id AS ClienteId,
                    c.Correo,
                    c.TipoDocumento,
                    c.NumeroDocumento,
                    c.DireccionFiscal,
                    c.CodigoUbigeo
                FROM dbo.Reservas r
                INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
                WHERE r.Id = @ReservaId AND s.NegocioId = @NegocioId;", cn);
            AddParam(cmd, "@ReservaId", reservaId, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
            await using var dr = await cmd.ExecuteReaderAsync();
            if (await dr.ReadAsync())
            {
                sedeId = dr.IsDBNull(0) ? null : dr.GetInt32(0);
                clienteId = dr.IsDBNull(1) ? null : dr.GetInt32(1);
                clienteCorreo = dr.IsDBNull(2) ? null : dr.GetString(2);
                clienteTipoDocumento = dr.IsDBNull(3) ? null : dr.GetString(3);
                clienteNumeroDocumento = dr.IsDBNull(4) ? null : dr.GetString(4);
                clienteDireccionFiscal = dr.IsDBNull(5) ? null : dr.GetString(5);
                clienteCodigoUbigeo = dr.IsDBNull(6) ? null : dr.GetString(6);
                if (!string.IsNullOrWhiteSpace(clienteCodigoUbigeo) && clienteCodigoUbigeo.Length >= 4)
                {
                    clienteCodigoDepartamento = clienteCodigoUbigeo[..2];
                    clienteCodigoProvincia = clienteCodigoUbigeo[..4];
                }
            }
        }

        var documentos = new List<SelectListItem>();
        if (sedeId.HasValue)
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmdDocs = new SqlCommand(@"
                SELECT DISTINCT
                    t.CodigoSunat AS Value,
                    CONCAT(t.Nombre, N' (', t.CodigoSunat, N')') AS Text
                FROM dbo.SedesSeriesDocumentoComprobante ss
                INNER JOIN dbo.NegociosSeriesDocumentoComprobante ns ON ns.Id = ss.NegocioSerieId
                INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ss.CodigoSunat
                INNER JOIN dbo.NegociosTiposDocumentoComprobante ntd ON ntd.NegocioId = ns.NegocioId AND ntd.CodigoSunat = ss.CodigoSunat
                WHERE ss.SedeId = @SedeId
                  AND ns.NegocioId = @NegocioId
                  AND ss.Activo = 1
                  AND ns.Activo = 1
                  AND ntd.Activo = 1
                  AND t.Activo = 1
                  AND t.Habilitado = 1
                ORDER BY Value;", cn);
            AddParam(cmdDocs, "@SedeId", sedeId.Value, SqlDbType.Int);
            AddParam(cmdDocs, "@NegocioId", negocioId, SqlDbType.Int);
            await using var drDocs = await cmdDocs.ExecuteReaderAsync();
            while (await drDocs.ReadAsync())
            {
                documentos.Add(new SelectListItem
                {
                    Value = drDocs.IsDBNull(0) ? string.Empty : drDocs.GetString(0),
                    Text = drDocs.IsDBNull(1) ? string.Empty : drDocs.GetString(1)
                });
            }
        }

        var codigoNormalizado = string.IsNullOrWhiteSpace(codigoSunat) ? null : codigoSunat.Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(codigoNormalizado))
        {
            codigoNormalizado = documentos
                .Select(x => x.Value?.Trim().ToUpperInvariant())
                .FirstOrDefault(v => v is "03" or "01" or "RI");
        }

        var series = new List<SelectListItem>();
        if (!string.IsNullOrWhiteSpace(codigoNormalizado) && sedeId.HasValue)
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmdSeries = new SqlCommand(@"
                SELECT
                    CONVERT(NVARCHAR(20), ns.Id) AS Value,
                    ns.Serie AS Text
                FROM dbo.SedesSeriesDocumentoComprobante ss
                INNER JOIN dbo.NegociosSeriesDocumentoComprobante ns ON ns.Id = ss.NegocioSerieId
                WHERE ss.SedeId = @SedeId
                  AND ss.CodigoSunat = @CodigoSunat
                  AND ss.Activo = 1
                  AND ns.NegocioId = @NegocioId
                  AND ns.Activo = 1
                ORDER BY ns.Serie;", cn);
            AddParam(cmdSeries, "@SedeId", sedeId.Value, SqlDbType.Int);
            AddParam(cmdSeries, "@CodigoSunat", codigoNormalizado, SqlDbType.NVarChar);
            AddParam(cmdSeries, "@NegocioId", negocioId, SqlDbType.Int);
            await using var drSeries = await cmdSeries.ExecuteReaderAsync();
            while (await drSeries.ReadAsync())
            {
                series.Add(new SelectListItem
                {
                    Value = drSeries.IsDBNull(0) ? string.Empty : drSeries.GetString(0),
                    Text = drSeries.IsDBNull(1) ? string.Empty : drSeries.GetString(1)
                });
            }
        }

        return new ComprobanteReservaContextoViewModel
        {
            ReservaId = pago.ReservaId,
            ReservaCodigo = pago.ReservaCodigo,
            Sede = pago.Sede,
            Espacio = pago.Espacio,
            Cliente = pago.Cliente,
            ClienteId = clienteId,
            ClienteCorreo = clienteCorreo,
            ClienteTipoDocumento = clienteTipoDocumento,
            ClienteNumeroDocumento = clienteNumeroDocumento,
            ClienteDireccionFiscal = clienteDireccionFiscal,
            ClienteCodigoUbigeo = clienteCodigoUbigeo,
            ClienteCodigoDepartamento = clienteCodigoDepartamento,
            ClienteCodigoProvincia = clienteCodigoProvincia,
            FechaReserva = pago.FechaReserva,
            HoraInicioReserva = pago.HoraInicioReserva,
            HoraFinReserva = pago.HoraFinReserva,
            TotalReserva = pago.TotalReserva,
            TotalPagado = pago.TotalPagado,
            SaldoPendiente = pago.SaldoPendiente,
            MonedaSimbolo = pago.MonedaSimbolo,
            PorcentajeIgvConfigurado = 18,
            PagosReserva = pago.Pagos.Select(x => new PagoPrevioItemViewModel
            {
                PagoId = x.PagoId,
                FechaPago = x.FechaPago,
                Monto = x.Monto,
                FormaPago = x.FormaPagoNombre,
                NumeroOperacion = x.NumeroOperacion,
                Observacion = x.Observacion
            }).ToList(),
            DocumentosDisponibles = documentos,
            SeriesDisponibles = series
        };
    }

    private static bool EsErrorNoEncontrado(string? mensaje)
    {
        if (string.IsNullOrWhiteSpace(mensaje)) return false;
        return mensaje.Contains("No se encontro", StringComparison.OrdinalIgnoreCase)
               || mensaje.Contains("no encontrada", StringComparison.OrdinalIgnoreCase)
               || mensaje.Contains("no encontrado", StringComparison.OrdinalIgnoreCase);
    }
}
