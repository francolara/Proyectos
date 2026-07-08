namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroElectronicoConsultaRequest
{
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public byte Mes { get; init; }
    public string LibroElectronico { get; init; } = PleLibroElectronicoCatalogo.LibroDiario51;
    public string Moneda { get; init; } = "PEN";
    public string Estado { get; init; } = "Todos";
    public DateOnly? FechaDesde { get; init; }
    public DateOnly? FechaHasta { get; init; }
}
