namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IReservationEmailNotificationService
{
    Task NotifyPublicReservationCreatedAsync(int? negocioId, int reservaId);
    Task NotifyReservationConfirmedIfAppliesAsync(int negocioId, int reservaId, int? estadoAnterior = null);
}
