using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<List<ReporteOcupacionItemViewModel>> ReportesOcupacionPorEspacioAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta)
    {
        var list = new List<ReporteOcupacionItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reportes_OcupacionPorEspacio", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        await using var dr = await cmd.ExecuteReaderAsync();
        while (await dr.ReadAsync())
        {
            list.Add(new ReporteOcupacionItemViewModel
            {
                Sede = dr.GetString(0),
                Espacio = dr.GetString(1),
                CantidadReservas = dr.GetInt32(2),
                HorasReservadas = dr.GetDecimal(3),
                MontoReservado = dr.GetDecimal(4),
                MontoCobrado = dr.GetDecimal(5)
            });
        }
        return list;
    }

    public async Task<List<ReporteIngresoDiaItemViewModel>> ReportesIngresosPorDiaAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta)
    {
        var list = new List<ReporteIngresoDiaItemViewModel>();
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Reportes_IngresosPorDia", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        AddParam(cmd, "@FechaDesde", fechaDesde.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
        AddParam(cmd, "@FechaHasta", fechaHasta.ToDateTime(TimeOnly.MinValue), SqlDbType.Date);
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
}
