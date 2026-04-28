using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IClubRegistrationNotificationService
{
    Task NotifyNewClubRegistrationAsync(AltaClubSolicitudFormViewModel request, string? requestCode = null);
    Task NotifyClubApprovalAsync(AltaClubItemViewModel request, int diasPrueba);
}
