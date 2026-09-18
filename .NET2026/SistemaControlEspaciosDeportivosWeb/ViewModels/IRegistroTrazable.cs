namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public interface IRegistroTrazable
{
    string? UsuarioCreacion { get; }
    DateTime? FechaRegistro { get; }
    string? UsuarioActualizacion { get; }
    DateTime? FechaActualizacion { get; }
}
