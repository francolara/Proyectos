using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<ReporteOcupacionItemViewModel>> ReportesOcupacionPorEspacioAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null)
    {
        var list = new List<ReporteOcupacionItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reportes_OcupacionPorEspacio", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReporteOcupacionItemViewModel
            {
                SedeId = dr.GetInt32(0),
                EspacioDeportivoId = dr.GetInt32(1),
                Sede = dr.GetString(2),
                Espacio = dr.GetString(3),
                CantidadReservas = dr.GetInt32(4),
                HorasReservadas = dr.GetDecimal(5),
                MontoReservado = dr.GetDecimal(6),
                MontoCobrado = dr.GetDecimal(7)
            });
        }
        return list;
    }

    public async Task<List<ReporteIngresoDiaItemViewModel>> ReportesIngresosPorDiaAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null)
    {
        var list = new List<ReporteIngresoDiaItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reportes_IngresosPorDia", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReporteIngresoDiaItemViewModel
            {
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(0)),
                CantidadReservas = dr.GetInt32(1),
                Ingresos = dr.GetDecimal(2)
            });
        }
        return list;
    }

    public async Task<List<ReporteReservaDiaItemViewModel>> ReportesReservasPorDiaAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null)
    {
        var list = new List<ReporteReservaDiaItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reportes_ReservasPorDia", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReporteReservaDiaItemViewModel
            {
                Fecha = DateOnly.FromDateTime(dr.GetDateTime(0)),
                CantidadReservas = dr.GetInt32(1),
                MontoReservado = dr.GetDecimal(2)
            });
        }
        return list;
    }

    public async Task<ReporteResumenOperativoViewModel> ReportesResumenOperativoAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null)
    {
        var vm = new ReporteResumenOperativoViewModel();

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reportes_ResumenOperativo", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (await dr.ReadAsync())
        {
            vm.TotalReservas = dr.IsDBNull(0) ? 0 : dr.GetInt32(0);
            vm.TotalPendientes = dr.IsDBNull(1) ? 0 : dr.GetInt32(1);
            vm.TotalConfirmadas = dr.IsDBNull(2) ? 0 : dr.GetInt32(2);
            vm.TotalPagadas = dr.IsDBNull(3) ? 0 : dr.GetInt32(3);
            vm.TotalCanceladas = dr.IsDBNull(4) ? 0 : dr.GetInt32(4);
            vm.TotalNoShow = dr.IsDBNull(5) ? 0 : dr.GetInt32(5);
            vm.MontoReservado = dr.IsDBNull(6) ? 0m : dr.GetDecimal(6);
            vm.MontoCobrado = dr.IsDBNull(7) ? 0m : dr.GetDecimal(7);
            vm.SaldoPendiente = dr.IsDBNull(8) ? 0m : dr.GetDecimal(8);
        }

        return vm;
    }

    public async Task<ReporteResumenCobranzaViewModel> ReportesResumenCobranzaAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null)
    {
        var vm = new ReporteResumenCobranzaViewModel();

        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reportes_ResumenCobranza", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@SedeId", sedeId, SqlDbType.Int);

        await using var dr = await cmd.ExecuteReaderAsync();
        if (await dr.ReadAsync())
        {
            vm.CantidadPagos = dr.IsDBNull(0) ? 0 : dr.GetInt32(0);
            vm.ReservasCobradas = dr.IsDBNull(1) ? 0 : dr.GetInt32(1);
            vm.MontoCobrado = dr.IsDBNull(2) ? 0m : dr.GetDecimal(2);
        }

        return vm;
    }
}
