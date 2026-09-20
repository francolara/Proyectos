namespace SistemaAdministrativoWeb.ViewModels;

/// <summary>
/// Datos de auditoria que se muestran al editar un registro existente.
/// </summary>
public interface IRegistroTrazable
{
    string? UsuarioCreacion { get; }
    DateTime? FechaRegistro { get; }
    string? UsuarioActualizacion { get; }
    DateTime? FechaActualizacion { get; }
}
