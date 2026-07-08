namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleFileNameService : IPleFileNameService
{
    public string ConstruirNombreArchivo(string ruc, short anio, byte mes, string libroElectronico, string moneda)
    {
        var rucTrabajo = (ruc ?? string.Empty).Trim();
        var periodo = PlePeriodoHelper.FormarPeriodo(anio, mes);
        var codigoLibro = PleLibroElectronicoCatalogo.ObtenerCodigoSunat(libroElectronico);
        var indicadorContenido = "1";
        var indicadorMoneda = string.Equals(moneda?.Trim(), "USD", StringComparison.OrdinalIgnoreCase) ? "2" : "1";
        var indicadorLibroElectronico = "11";

        return $"LE{rucTrabajo}{periodo}{codigoLibro}00{indicadorContenido}{indicadorMoneda}{indicadorLibroElectronico}.txt";
    }
}
