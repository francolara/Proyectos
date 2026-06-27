namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class DetraccionSunatDto
{
    public int IdDetraccionSunat { get; init; }
    public string CodigoSunat { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public decimal Porcentaje { get; init; }
    public bool Estado { get; init; }
}
