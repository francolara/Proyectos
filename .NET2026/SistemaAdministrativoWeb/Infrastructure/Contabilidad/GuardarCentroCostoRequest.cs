namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarCentroCostoRequest
{
    public int? IdCentroCosto { get; init; }
    public int IdEmpresa { get; init; }
    public string CodigoCentroCosto { get; init; } = string.Empty;
    public string NombreCentroCosto { get; init; } = string.Empty;
    public bool Estado { get; init; }
}
