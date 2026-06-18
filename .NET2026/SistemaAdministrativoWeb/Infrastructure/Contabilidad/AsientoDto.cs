namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoDto : AsientoResumenDto
{
    public List<AsientoDetalleDto> Detalles { get; init; } = [];
}
