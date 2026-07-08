namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoPercepcionDto
{
    public int IdTipoPercepcion { get; init; }
    public string Codigo { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public decimal Porcentaje { get; init; }
    public bool Estado { get; init; }
}
