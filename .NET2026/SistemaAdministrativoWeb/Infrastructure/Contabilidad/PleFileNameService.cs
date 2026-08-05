namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

// Firma: FRANCO LARA - 04/08/2026 | Corrige las posiciones de los indicadores de operaciones, contenido, moneda y generación en los nombres PLE.
public sealed class PleFileNameService : IPleFileNameService
{
    public string ConstruirNombreArchivo(string ruc, short anio, byte mes, string libroElectronico, string moneda, bool tieneContenido = true)
    {
        var rucTrabajo = (ruc ?? string.Empty).Trim();
        var periodo = PlePeriodoHelper.FormarPeriodo(anio, mes);
        var codigoLibro = PleLibroElectronicoCatalogo.ObtenerCodigoSunat(libroElectronico);
        var indicadorOperaciones = "1";
        var indicadorContenido = tieneContenido ? "1" : "0";
        var indicadorMoneda = "1";
        var indicadorLibroElectronico = "1";

        return $"LE{rucTrabajo}{periodo}{codigoLibro}00{indicadorOperaciones}{indicadorContenido}{indicadorMoneda}{indicadorLibroElectronico}.txt";
    }
}
