namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public static class PlePeriodoHelper
{
    public static string FormarPeriodo(short anio, byte mes)
    {
        return $"{anio:0000}{mes:00}00";
    }

    public static string FormarPeriodoContable(short anio, byte mes)
    {
        return $"{anio:0000}{mes:00}";
    }
}
