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

    public async Task<List<PagoItemViewModel>> PagosListarAsync(int negocioId, int? sedeId = null)
    {
        var list = new List<PagoItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync()) list.Add(new PagoItemViewModel { Id = dr.GetInt32(0), ReservaId = dr.GetInt32(1), FechaPago = dr.GetDateTime(2), Monto = dr.GetDecimal(3), FormaPago = dr.GetString(4) });
        return list;
    }

    public async Task<PagoFormViewModel?> PagosObtenerAsync(int negocioId, int id)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_ObtenerPorId", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return null;
        return new PagoFormViewModel
        {
            Id = dr.GetInt32(0),
            ReservaId = dr.GetInt32(1),
            FechaPago = dr.GetDateTime(2),
            Monto = dr.GetDecimal(3),
            FormaPagoId = dr.GetInt32(4),
            NumeroOperacion = dr.IsDBNull(5) ? null : dr.GetString(5),
            Observacion = dr.IsDBNull(6) ? null : dr.GetString(6),
            NegocioId = negocioId
        };
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

    public async Task<bool> PagosActualizarAsync(PagoFormViewModel model, string usuario)
    {
        try
        {
            await using var cn = CreateConnection();
            await cn.OpenAsync();
            await using var cmd = new SqlCommand("Sp_Pagos_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
            AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
            AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
            AddParam(cmd, "@ReservaId", model.ReservaId, SqlDbType.Int);
            AddParam(cmd, "@FechaPago", model.FechaPago, SqlDbType.DateTime2);
            AddParam(cmd, "@Monto", model.Monto, SqlDbType.Decimal);
            AddParam(cmd, "@FormaPago", model.FormaPagoId, SqlDbType.Int);
            AddParam(cmd, "@NumeroOperacion", model.NumeroOperacion, SqlDbType.NVarChar);
            AddParam(cmd, "@Observacion", model.Observacion, SqlDbType.NVarChar);
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

    public Task<List<SelectListItem>> PagosComboReservasAsync(int negocioId, int? sedeId = null) => ComboAsync("Sp_Combos_ReservasPorNegocio", ("@NegocioId", (object?)negocioId, SqlDbType.Int), ("@SedeId", sedeId, SqlDbType.Int));
    public Task<List<SelectListItem>> PagosComboFormasPagoAsync(int negocioId) => ComboAsync("Sp_Combos_FormasPago", ("@NegocioId", negocioId, SqlDbType.Int));

    public async Task<List<ComprobanteItemViewModel>> ComprobantesListarAsync(int negocioId, int? sedeId = null)
    {
        var list = new List<ComprobanteItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
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
                Estado = dr.GetString(6)
            });
        }
        return list;
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
        AddParam(cmd, "@Serie", model.Serie, SqlDbType.NVarChar);
        AddParam(cmd, "@Numero", model.Numero, SqlDbType.Int);
        AddParam(cmd, "@FechaEmision", model.FechaEmision, SqlDbType.DateTime2);
        AddParam(cmd, "@TipoMoneda", (int)model.TipoMoneda, SqlDbType.Int);
        AddParam(cmd, "@SubTotal", model.SubTotal, SqlDbType.Decimal);
        AddParam(cmd, "@Igv", model.Igv, SqlDbType.Decimal);
        AddParam(cmd, "@Total", model.Total, SqlDbType.Decimal);
        AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
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
            AddParam(cmd, "@ReservaId", model.ReservaId, SqlDbType.Int);
            AddParam(cmd, "@TipoComprobante", (int)model.TipoComprobante, SqlDbType.Int);
            AddParam(cmd, "@Serie", model.Serie, SqlDbType.NVarChar);
            AddParam(cmd, "@Numero", model.Numero, SqlDbType.Int);
            AddParam(cmd, "@FechaEmision", model.FechaEmision, SqlDbType.DateTime2);
            AddParam(cmd, "@TipoMoneda", (int)model.TipoMoneda, SqlDbType.Int);
            AddParam(cmd, "@SubTotal", model.SubTotal, SqlDbType.Decimal);
            AddParam(cmd, "@Igv", model.Igv, SqlDbType.Decimal);
            AddParam(cmd, "@Total", model.Total, SqlDbType.Decimal);
            AddParam(cmd, "@Estado", (int)model.Estado, SqlDbType.Int);
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

    private static bool EsErrorNoEncontrado(string? mensaje)
    {
        if (string.IsNullOrWhiteSpace(mensaje)) return false;
        return mensaje.Contains("No se encontro", StringComparison.OrdinalIgnoreCase)
               || mensaje.Contains("no encontrada", StringComparison.OrdinalIgnoreCase)
               || mensaje.Contains("no encontrado", StringComparison.OrdinalIgnoreCase);
    }
}
