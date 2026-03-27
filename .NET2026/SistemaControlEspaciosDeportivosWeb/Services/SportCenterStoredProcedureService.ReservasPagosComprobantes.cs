using System.Data;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<ReservaItemViewModel>> ReservasListarAsync(int negocioId, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int? sedeId = null, int? espacioDeportivoId = null, int? estado = null)
    {
        var list = new List<ReservaItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta?.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        AddParam(cmd, "@EspacioDeportivoId", espacioDeportivoId, SqlDbType.Int);
        AddParam(cmd, "@Estado", estado, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReservaItemViewModel
            {
                Id = dr.GetInt32(0),
                Cliente = dr.GetString(1),
                Espacio = dr.GetString(2),
                Sede = dr.GetString(3),
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(4)),
                HoraInicio = TimeOnly.FromTimeSpan(dr.GetTimeSpan(5)),
                HoraFin = TimeOnly.FromTimeSpan(dr.GetTimeSpan(6)),
                Total = dr.GetDecimal(7),
                Estado = dr.GetString(8)
            });
        }
        return list;
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
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> ReservasActualizarAsync(ReservaFormViewModel model, string usuario)
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
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> ReservasEliminarAsync(int negocioId, int id, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> ReservasCambiarEstadoRapidoAsync(int negocioId, int id, int nuevoEstado, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reservas_CambiarEstadoRapido", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@NuevoEstado", nuevoEstado, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
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
                Sede = dr.IsDBNull(10) ? string.Empty : dr.GetString(10)
            });
        }
        return list;
    }

    public async Task<bool> ReservasMoverAsync(int negocioId, int id, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin, string usuario)
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
        return await cmd.ExecuteNonQueryAsync() > 0;
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
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public Task<List<SelectListItem>> ReservasComboEspaciosAsync(int negocioId) => ComboAsync("Sp_Combos_EspaciosPorNegocio", ("@NegocioId", (object?)negocioId, SqlDbType.Int));
    public Task<List<SelectListItem>> ReservasComboClientesAsync(int negocioId) => ComboAsync("Sp_Combos_Clientes", ("@NegocioId", (object?)negocioId, SqlDbType.Int));

    public async Task<List<PagoItemViewModel>> PagosListarAsync(int negocioId)
    {
        var list = new List<PagoItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
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
            FormaPago = (FormaPago)dr.GetInt32(4),
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
        AddParam(cmd, "@FormaPago", (int)model.FormaPago, SqlDbType.Int);
        AddParam(cmd, "@NumeroOperacion", model.NumeroOperacion, SqlDbType.NVarChar);
        AddParam(cmd, "@Observacion", model.Observacion, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return Convert.ToInt32(await cmd.ExecuteScalarAsync());
    }

    public async Task<bool> PagosActualizarAsync(PagoFormViewModel model, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_Actualizar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@Id", model.Id, SqlDbType.Int);
        AddParam(cmd, "@NegocioId", model.NegocioId, SqlDbType.Int);
        AddParam(cmd, "@ReservaId", model.ReservaId, SqlDbType.Int);
        AddParam(cmd, "@FechaPago", model.FechaPago, SqlDbType.DateTime2);
        AddParam(cmd, "@Monto", model.Monto, SqlDbType.Decimal);
        AddParam(cmd, "@FormaPago", (int)model.FormaPago, SqlDbType.Int);
        AddParam(cmd, "@NumeroOperacion", model.NumeroOperacion, SqlDbType.NVarChar);
        AddParam(cmd, "@Observacion", model.Observacion, SqlDbType.NVarChar);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> PagosEliminarAsync(int negocioId, int id, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Pagos_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public Task<List<SelectListItem>> PagosComboReservasAsync(int negocioId) => ComboAsync("Sp_Combos_ReservasPorNegocio", ("@NegocioId", (object?)negocioId, SqlDbType.Int));

    public async Task<List<ComprobanteItemViewModel>> ComprobantesListarAsync(int negocioId)
    {
        var list = new List<ComprobanteItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_Listar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
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
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> ComprobantesEliminarAsync(int negocioId, int id, string usuario)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_Eliminar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@Id", id, SqlDbType.Int);
        AddParam(cmd, "@Usuario", usuario, SqlDbType.NVarChar);
        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public Task<List<SelectListItem>> ComprobantesComboReservasAsync(int negocioId) => ComboAsync("Sp_Combos_ReservasPorNegocio", ("@NegocioId", (object?)negocioId, SqlDbType.Int));
}
