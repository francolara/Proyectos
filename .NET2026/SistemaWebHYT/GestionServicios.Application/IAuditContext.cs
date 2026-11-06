namespace GestionServicios.Application;

// El host obtiene la identidad del usuario autenticado, nunca del formulario.
public interface IAuditContext
{
    string UserId { get; }
    Guid CorrelationId { get; }
}
