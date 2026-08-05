namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPleFileNameService
{
    string ConstruirNombreArchivo(string ruc, short anio, byte mes, string libroElectronico, string moneda, bool tieneContenido = true);
}
